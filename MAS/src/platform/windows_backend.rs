//! Real Windows Software Protection Platform backend, via WMI/COM.
//!
//! Compiled only under `cfg(all(windows, feature = "winapi"))`. It drives the
//! same WMI classes on `ROOT\CIMV2` that MAS drives through PowerShell/`slmgr`:
//! `SoftwareLicensingService` / `SoftwareLicensingProduct` and the Office
//! `OfficeSoftwareProtection*` equivalents.
//!
//! All `unsafe`, all COM lifetime handling and every raw `VARIANT` access are
//! contained in this file; the rest of the crate sees only the safe [`Spp`]
//! trait (constraint 3: FFI behind a sound safe API).
//!
//! Written against `windows` 0.58. It is intentionally **not compiled or tested
//! on the porting host** (a Linux box has no Windows target); treat signatures
//! as knowledge-derived and expect to shake out minor API drift on a real
//! Windows build. The logic — WQL queries, method names, parameter names and
//! the `ReturnValue` HRESULT handling — mirrors the SPP contract exactly.

#![allow(clippy::missing_safety_doc)]

use crate::error::{Error, Result};
use crate::model::{LicenseStatus, Product};
use crate::platform::{LicenseInfo, Spp};
use std::path::Path;

use windows::core::{Interface, BSTR, HSTRING, PCWSTR};
use windows::Win32::Foundation::HANDLE;
use windows::Win32::Security::{
    GetTokenInformation, TokenElevation, TOKEN_ELEVATION, TOKEN_QUERY,
};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoInitializeSecurity, CLSCTX_INPROC_SERVER,
    COINIT_MULTITHREADED, EOAC_NONE, RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE,
};
use windows::Win32::System::Threading::{GetCurrentProcess, OpenProcessToken};
use windows::Win32::System::Variant::{
    VariantClear, VARIANT, VT_BSTR, VT_I4, VT_UI4,
};
use windows::Win32::System::Wmi::{
    IWbemClassObject, IWbemLocator, IWbemServices, WbemLocator, WBEM_FLAG_FORWARD_ONLY,
    WBEM_FLAG_RETURN_IMMEDIATELY, WBEM_INFINITE,
};

/// WMI-backed SPP. Holds a live `IWbemServices` connection to `ROOT\CIMV2`.
pub struct WindowsSpp {
    services: IWbemServices,
}

// ---- VARIANT helpers (the only place raw union fields are touched) ----------

/// Build a `VARIANT` holding a BSTR string.
fn variant_bstr(s: &str) -> VARIANT {
    // `VARIANT::from(&BSTR)` sets vt = VT_BSTR and owns the string; using the
    // crate's conversion keeps us off the raw union for construction.
    VARIANT::from(BSTR::from(s))
}

/// Build a `VARIANT` holding an unsigned 32-bit value.
fn variant_u32(v: u32) -> VARIANT {
    VARIANT::from(v as i32)
}

/// Extract a string from a `VARIANT`, whatever numeric/bstr shape it holds.
///
/// # Safety
/// `var` must be a valid, initialised VARIANT produced by a WMI `Get`.
unsafe fn variant_to_string(var: &VARIANT) -> Option<String> {
    let vt = var.Anonymous.Anonymous.vt;
    let val = &var.Anonymous.Anonymous.Anonymous;
    let out = match vt {
        VT_BSTR => {
            let b = &val.bstrVal;
            if b.is_empty() {
                return None;
            }
            b.to_string()
        }
        VT_I4 => val.lVal.to_string(),
        VT_UI4 => val.ulVal.to_string(),
        _ => return None,
    };
    Some(out)
}

/// Extract a u32 from a numeric `VARIANT`.
///
/// # Safety
/// `var` must be a valid, initialised VARIANT.
unsafe fn variant_to_u32(var: &VARIANT) -> Option<u32> {
    let vt = var.Anonymous.Anonymous.vt;
    let val = &var.Anonymous.Anonymous.Anonymous;
    match vt {
        VT_UI4 => Some(val.ulVal),
        VT_I4 => Some(val.lVal as u32),
        VT_BSTR => val.bstrVal.to_string().parse().ok(),
        _ => None,
    }
}

// ---- WMI object property access --------------------------------------------

/// Read a named property off a WMI object as a string.
fn get_string(obj: &IWbemClassObject, name: &str) -> Option<String> {
    let wname = HSTRING::from(name);
    let mut var = VARIANT::default();
    unsafe {
        obj.Get(PCWSTR(wname.as_ptr()), 0, &mut var, None, None).ok()?;
        let out = variant_to_string(&var);
        let _ = VariantClear(&mut var);
        out
    }
}

/// Read a named property off a WMI object as a u32.
fn get_u32(obj: &IWbemClassObject, name: &str) -> Option<u32> {
    let wname = HSTRING::from(name);
    let mut var = VARIANT::default();
    unsafe {
        obj.Get(PCWSTR(wname.as_ptr()), 0, &mut var, None, None).ok()?;
        let out = variant_to_u32(&var);
        let _ = VariantClear(&mut var);
        out
    }
}

impl WindowsSpp {
    /// Initialise COM and connect to the WMI licensing namespace.
    pub fn new() -> Self {
        Self::connect().unwrap_or_else(|e| panic!("WMI init failed: {e}"))
    }

    fn connect() -> Result<Self> {
        unsafe {
            // MTA — the process is a short-lived console tool.
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .ok()
                .map_err(|e| Error::winapi("CoInitializeEx", e.code().0))?;

            // Default process-wide security for WMI. Ignore RPC_E_TOO_LATE if a
            // host already initialised security.
            let _ = CoInitializeSecurity(
                None,
                -1,
                None,
                None,
                RPC_C_AUTHN_LEVEL_DEFAULT,
                RPC_C_IMP_LEVEL_IMPERSONATE,
                None,
                EOAC_NONE,
                None,
            );

            let locator: IWbemLocator =
                CoCreateInstance(&WbemLocator, None, CLSCTX_INPROC_SERVER)
                    .map_err(|e| Error::winapi("create IWbemLocator", e.code().0))?;

            let services = locator
                .ConnectServer(
                    &BSTR::from("ROOT\\CIMV2"),
                    &BSTR::new(),
                    &BSTR::new(),
                    &BSTR::new(),
                    0,
                    &BSTR::new(),
                    None,
                )
                .map_err(|e| Error::winapi("WMI ConnectServer", e.code().0))?;

            Ok(WindowsSpp { services })
        }
    }

    /// Run a WQL query and collect the result objects.
    fn query(&self, wql: &str) -> Result<Vec<IWbemClassObject>> {
        unsafe {
            let enumerator = self
                .services
                .ExecQuery(
                    &BSTR::from("WQL"),
                    &BSTR::from(wql),
                    WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
                    None,
                )
                .map_err(|e| Error::winapi("WMI ExecQuery", e.code().0))?;

            let mut out = Vec::new();
            loop {
                let mut row: [Option<IWbemClassObject>; 1] = [None];
                let mut returned = 0u32;
                // Next returns S_FALSE (and returned==0) at end of enumeration.
                let hr = enumerator.Next(WBEM_INFINITE.0, &mut row, &mut returned);
                if returned == 0 || hr.is_err() {
                    break;
                }
                if let Some(obj) = row[0].take() {
                    out.push(obj);
                }
            }
            Ok(out)
        }
    }

    /// Fetch the WMI `__PATH` of the single service instance of `class`.
    fn service_path(&self, class: &str) -> Result<String> {
        let rows = self.query(&format!("SELECT __PATH FROM {class}"))?;
        rows.first()
            .and_then(|o| get_string(o, "__PATH"))
            .ok_or_else(|| Error::winapi_msg(format!("no {class} instance")))
    }

    /// Invoke a WMI method on `object_path` (an instance or service `__PATH`),
    /// setting the given string/u32 parameters, and return its `ReturnValue`
    /// (the SPP HRESULT: 0 = success).
    fn call_method(
        &self,
        object_path: &str,
        class: &str,
        method: &str,
        params: &[(&str, ParamValue)],
    ) -> Result<i32> {
        unsafe {
            // Build the in-parameters instance from the class' method signature.
            let mut class_obj: Option<IWbemClassObject> = None;
            self.services
                .GetObject(
                    &BSTR::from(class),
                    Default::default(),
                    None,
                    Some(&mut class_obj),
                    None,
                )
                .map_err(|e| Error::winapi("WMI GetObject(class)", e.code().0))?;
            let class_obj =
                class_obj.ok_or_else(|| Error::winapi_msg(format!("class {class} not found")))?;

            let in_params = if params.is_empty() {
                None
            } else {
                let wmethod = HSTRING::from(method);
                let mut in_sig: Option<IWbemClassObject> = None;
                let mut out_sig: Option<IWbemClassObject> = None;
                class_obj
                    .GetMethod(PCWSTR(wmethod.as_ptr()), 0, &mut in_sig, &mut out_sig)
                    .map_err(|e| Error::winapi("WMI GetMethod", e.code().0))?;
                let in_sig = in_sig
                    .ok_or_else(|| Error::winapi_msg(format!("method {method} has no in-params")))?;
                let inst = in_sig
                    .SpawnInstance(0)
                    .map_err(|e| Error::winapi("WMI SpawnInstance", e.code().0))?;
                for (pname, pval) in params {
                    let wname = HSTRING::from(*pname);
                    let mut var = match pval {
                        ParamValue::Str(s) => variant_bstr(s),
                        ParamValue::U32(v) => variant_u32(*v),
                    };
                    inst.Put(PCWSTR(wname.as_ptr()), 0, &var, 0)
                        .map_err(|e| Error::winapi("WMI Put(param)", e.code().0))?;
                    let _ = VariantClear(&mut var);
                }
                Some(inst)
            };

            let mut out_params: Option<IWbemClassObject> = None;
            self.services
                .ExecMethod(
                    &BSTR::from(object_path),
                    &BSTR::from(method),
                    Default::default(),
                    None,
                    in_params.as_ref(),
                    Some(&mut out_params),
                    None,
                )
                .map_err(|e| Error::winapi(format!("WMI ExecMethod {method}"), e.code().0))?;

            // ReturnValue is the operation's own HRESULT; absent ⇒ success.
            let ret = out_params
                .as_ref()
                .and_then(|o| get_u32(o, "ReturnValue"))
                .unwrap_or(0) as i32;
            Ok(ret)
        }
    }

    /// Wrap [`call_method`] turning a non-zero `ReturnValue` into an error.
    fn call_checked(
        &self,
        object_path: &str,
        class: &str,
        method: &str,
        params: &[(&str, ParamValue)],
    ) -> Result<()> {
        match self.call_method(object_path, class, method, params)? {
            0 => Ok(()),
            hr => Err(Error::winapi(format!("{class}.{method}"), hr)),
        }
    }

    /// `__PATH` of a specific product row identified by its activation ID.
    fn product_path(&self, product: Product, activation_id: &str) -> Result<String> {
        let wql = format!(
            "SELECT __PATH FROM {class} WHERE ID='{id}'",
            class = product.product_class(),
            id = escape(activation_id),
        );
        self.query(&wql)?
            .first()
            .and_then(|o| get_string(o, "__PATH"))
            .ok_or_else(|| Error::winapi_msg(format!("product {activation_id} not found")))
    }
}

/// Value kinds accepted by [`WindowsSpp::call_method`].
enum ParamValue {
    Str(String),
    U32(u32),
}

/// Minimal WQL string-literal escaping (single quote → doubled).
fn escape(s: &str) -> String {
    s.replace('\'', "''")
}

impl Spp for WindowsSpp {
    fn installed_products(&self, product: Product) -> Result<Vec<LicenseInfo>> {
        let wql = format!(
            "SELECT ID, Name, Description, PartialProductKey, LicenseStatus, LicenseFamily, \
             GracePeriodRemaining FROM {class} WHERE ApplicationID='{app}' \
             AND PartialProductKey IS NOT NULL AND LicenseDependsOn IS NULL",
            class = product.product_class(),
            app = product.application_id(),
        );
        let mut out = Vec::new();
        for obj in self.query(&wql)? {
            let status_raw = get_u32(&obj, "LicenseStatus").unwrap_or(0);
            out.push(LicenseInfo {
                activation_id: get_string(&obj, "ID").unwrap_or_default(),
                name: get_string(&obj, "Name").unwrap_or_default(),
                description: get_string(&obj, "Description").unwrap_or_default(),
                partial_product_key: get_string(&obj, "PartialProductKey"),
                status: LicenseStatus::from_wmi(status_raw).unwrap_or(LicenseStatus::Unlicensed),
                license_family: get_string(&obj, "LicenseFamily"),
                grace_minutes: get_u32(&obj, "GracePeriodRemaining"),
            });
        }
        Ok(out)
    }

    fn install_product_key(&self, product: Product, key: &str) -> Result<()> {
        let class = product.service_class();
        let path = self.service_path(class)?;
        self.call_checked(
            &path,
            class,
            "InstallProductKey",
            &[("ProductKey", ParamValue::Str(key.to_string()))],
        )
    }

    fn uninstall_product_key(&self, product: Product, activation_id: &str) -> Result<()> {
        let path = self.product_path(product, activation_id)?;
        self.call_checked(&path, product.product_class(), "UninstallProductKey", &[])
    }

    fn set_kms_host(&self, product: Product, host: &str, port: u16) -> Result<()> {
        let class = product.service_class();
        let path = self.service_path(class)?;
        self.call_checked(
            &path,
            class,
            "SetKeyManagementServiceMachine",
            &[("MachineName", ParamValue::Str(host.to_string()))],
        )?;
        self.call_checked(
            &path,
            class,
            "SetKeyManagementServicePort",
            &[("PortNumber", ParamValue::U32(port as u32))],
        )
    }

    fn clear_kms_host(&self, product: Product) -> Result<()> {
        let class = product.service_class();
        let path = self.service_path(class)?;
        self.call_checked(&path, class, "ClearKeyManagementServiceMachine", &[])?;
        self.call_checked(&path, class, "ClearKeyManagementServicePort", &[])
    }

    fn activate(&self, product: Product, activation_id: &str) -> Result<()> {
        let path = self.product_path(product, activation_id)?;
        self.call_checked(&path, product.product_class(), "Activate", &[])
    }

    fn install_license(&self, xml_path: &Path) -> Result<()> {
        let xml = std::fs::read_to_string(xml_path)?;
        let class = Product::Windows.service_class();
        let path = self.service_path(class)?;
        self.call_checked(
            &path,
            class,
            "InstallLicense",
            &[("License", ParamValue::Str(xml))],
        )
    }

    fn windows_build(&self) -> Result<u32> {
        let rows = self.query("SELECT BuildNumber FROM Win32_OperatingSystem")?;
        rows.first()
            .and_then(|o| get_string(o, "BuildNumber"))
            .and_then(|s| s.trim().parse().ok())
            .ok_or_else(|| Error::winapi_msg("could not read Windows build"))
    }

    fn windows_edition(&self) -> Result<String> {
        // Best-effort from the OS caption, e.g. "Microsoft Windows 11 Pro".
        let rows = self.query("SELECT Caption FROM Win32_OperatingSystem")?;
        Ok(rows
            .first()
            .and_then(|o| get_string(o, "Caption"))
            .unwrap_or_else(|| "Unknown".to_string()))
    }

    fn is_elevated(&self) -> bool {
        unsafe {
            let process = GetCurrentProcess();
            let mut token = HANDLE::default();
            if OpenProcessToken(process, TOKEN_QUERY, &mut token).is_err() {
                return false;
            }
            let mut elevation = TOKEN_ELEVATION::default();
            let mut size = 0u32;
            let ok = GetTokenInformation(
                token,
                TokenElevation,
                Some(&mut elevation as *mut _ as *mut _),
                std::mem::size_of::<TOKEN_ELEVATION>() as u32,
                &mut size,
            )
            .is_ok();
            let _ = windows::Win32::Foundation::CloseHandle(token);
            ok && elevation.TokenIsElevated != 0
        }
    }

    // --- Digital-license / ClipSVC operations --------------------------------

    fn write_genuine_ticket(&self, xml: &[u8]) -> Result<()> {
        let dir = clipsvc_dir().join("GenuineTicket");
        std::fs::create_dir_all(&dir)?;
        // MAS writes `GenuineTicket` then copies it to `GenuineTicket.xml`;
        // ClipSVC consumes the `.xml`.
        std::fs::write(dir.join("GenuineTicket"), xml)?;
        std::fs::write(dir.join("GenuineTicket.xml"), xml)?;
        Ok(())
    }

    fn restart_service(&self, name: &str) -> Result<()> {
        // Faithful to the script's `Restart-Service` (PowerShell handles
        // dependent services; `net`/`sc` do not).
        run_tool(
            "powershell",
            &[
                "-NoProfile",
                "-Command",
                &format!("Restart-Service -Name '{name}' -Force"),
            ],
        )
    }

    fn run_clipup(&self, args: &[&str]) -> Result<()> {
        // ClipUp.exe lives in System32, which is on PATH for an elevated shell.
        run_tool("clipup", args)
    }

    fn clip_tokens_present(&self) -> bool {
        clipsvc_dir().join("tokens.dat").exists()
    }

    fn pin_kms38(&self, activation_id: &str) -> Result<()> {
        // Per-activation-ID key under the Windows SPP registry root.
        let key = format!(
            r"HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\{}\{}",
            Product::Windows.application_id(),
            activation_id
        );
        run_tool(
            "reg",
            &[
                "add", key.as_str(), "/v", "KeyManagementServiceName", "/t", "REG_SZ", "/d",
                "127.0.0.2", "/f",
            ],
        )?;
        run_tool(
            "reg",
            &[
                "add", key.as_str(), "/v", "KeyManagementServicePort", "/t", "REG_SZ", "/d", "1688",
                "/f",
            ],
        )
    }
}

/// `%ProgramData%\Microsoft\Windows\ClipSVC`.
fn clipsvc_dir() -> std::path::PathBuf {
    let pd = std::env::var("ProgramData").unwrap_or_else(|_| r"C:\ProgramData".to_string());
    std::path::Path::new(&pd).join(r"Microsoft\Windows\ClipSVC")
}

/// Spawn an external tool and map a non-zero exit to an [`Error::ExternalTool`].
fn run_tool(tool: &str, args: &[&str]) -> Result<()> {
    let status = std::process::Command::new(tool)
        .args(args)
        .status()
        .map_err(|e| Error::ExternalTool {
            tool: tool.to_string(),
            detail: e.to_string(),
        })?;
    if status.success() {
        Ok(())
    } else {
        Err(Error::ExternalTool {
            tool: tool.to_string(),
            detail: format!("exited with {status}"),
        })
    }
}

impl Drop for WindowsSpp {
    fn drop(&mut self) {
        // Balance CoInitializeEx. Safe: called once per successful init.
        unsafe { windows::Win32::System::Com::CoUninitialize() };
    }
}

impl Default for WindowsSpp {
    fn default() -> Self {
        Self::new()
    }
}
