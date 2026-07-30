//! Thin entry point — all logic lives in the library so it can be unit-tested.

fn main() -> std::process::ExitCode {
    match mas::run() {
        Ok(()) => std::process::ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("mas: {e}");
            std::process::ExitCode::FAILURE
        }
    }
}
