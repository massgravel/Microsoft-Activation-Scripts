#include <utility>
#include "mem_scan.hpp"
#include "mem_iter.hpp"
#include "mem_util.hpp"


template <typename ...Args>
__forceinline void output_to_console( const char* str, Args&&... args )
{
	DbgPrintEx( 77, 0, str, std::forward<Args>( args )... );
}

__forceinline void output_appended( const char* str )
{
	output_to_console( "[!] watermark_disabler: %s\n", str );
}

NTSTATUS driver_entry( )
{
	output_appended( "loaded" );

	/* we have to attach to csrss, or any process with win32k mapped into it, because win32k is not mapped in system modules */
	const auto csrss_process = impl::search_for_process( "csrss.exe" );

	if ( !csrss_process )
	{
		output_appended( "failed to find csrss.exe" );
		return STATUS_UNSUCCESSFUL;
	}

	impl::unique_attachment csrss_attach( csrss_process );

	output_appended( "attached to csrss" );

	const auto win32kfull_info = impl::search_for_module( "win32kfull.sys" );

	if ( !win32kfull_info )
	{
		output_appended( "failed to find the win32kfull.sys module" );
		return STATUS_UNSUCCESSFUL;
	}

	output_to_console( "[!] watermark_disabler: win32kfull.sys $ 0x%p\n", win32kfull_info->image_base );

	const auto gpsi_instruction = impl::search_for_signature( win32kfull_info, "\x48\x8b\x0d\x00\x00\x00\x00\x48\x8b\x05\x00\x00\x00\x00\x0f\xba\x30\x0c", "xxx????xxx????xxxx" );

	if ( !gpsi_instruction )
	{
		output_appended( "failed to find gpsi, signature outdated?" );
		return STATUS_UNSUCCESSFUL;
	}

	const auto gpsi = *reinterpret_cast< std::uint64_t* >( impl::resolve_mov( gpsi_instruction ) );

	if ( !gpsi )
	{
		output_appended( "gpsi is somehow nullptr" );
		return STATUS_UNSUCCESSFUL;
	}

	output_to_console( "[!] watermark_disabler: gpsi $ 0x%p\n", gpsi );

	*reinterpret_cast< std::uint32_t* >( gpsi + 0x874 ) = 0;

	output_appended( "watermark disabled" );

	return STATUS_UNSUCCESSFUL;
}
