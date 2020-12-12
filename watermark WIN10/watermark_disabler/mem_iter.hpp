#pragma once
#include "mem_defs.hpp"
#include "util_raii.hpp"

namespace impl
{
	extern "C" NTSYSAPI NTSTATUS NTAPI ZwQuerySystemInformation( ULONG, PVOID, ULONG, PULONG );
	extern "C" NTSYSAPI PIMAGE_NT_HEADERS NTAPI RtlImageNtHeader( PVOID );

	nt::rtl_module_info* search_for_module( const char* module_name )
	{
		/* allocate a pool with 0x2000 bytes because we don't know how big the module list is */
		auto needed_bytes = 8192ul;
		impl::pool buffer_pool( ExAllocatePoolWithTag( PagedPool, needed_bytes, 'udoM' ) );

		if ( !buffer_pool.get( ) )
			return nullptr;

		auto current_status = ZwQuerySystemInformation( 11, buffer_pool.get( ), needed_bytes, &needed_bytes );

		/* keep allocating until the function returns STATUS_SUCCESS */
		while ( current_status == STATUS_INFO_LENGTH_MISMATCH )
		{
			buffer_pool.reset( ExAllocatePoolWithTag( PagedPool, needed_bytes, 'udoM' ) );

			if ( !buffer_pool )
				return nullptr;

			current_status = ZwQuerySystemInformation( 11, buffer_pool.get( ), needed_bytes, &needed_bytes );
		}

		if ( !NT_SUCCESS( current_status ) )
			return nullptr;

		const auto current_modules = static_cast< nt::rtl_modules* >( buffer_pool.get( ) );

		if ( !current_modules )
			return nullptr;

		/* loop the module list, and find the needed module */
		for ( auto i = 0u; i < current_modules->count; i++ )
		{
			const auto current_module = &current_modules->modules[ i ];

			if ( !current_module )
				continue;

			/* file_name_offset is the offset from full_path to the actual file's name, instead of file path */
			const auto file_name = reinterpret_cast< const char* >( current_module->file_name_offset + current_module->full_path );

			if ( std::strcmp( file_name, module_name ) != 0 )
				continue;

			return current_module;
		}

		return nullptr;
	}

	nt::image_section_header* search_for_section( const nt::rtl_module_info* module, const char* section_name )
	{
		if ( !module )
			return nullptr;

		const auto nt_header = reinterpret_cast< nt::image_nt_headers* >( RtlImageNtHeader( module->image_base ) );
		const auto section_header = reinterpret_cast< nt::image_section_header* >( nt_header + 1 );

		for ( auto i = 0u; i < nt_header->file_header.number_of_sections; i++ )
		{
			const auto current_section = &section_header[ i ];
			const auto current_section_name = reinterpret_cast< const char* >( current_section->name );

			if ( !std::strcmp( current_section_name, section_name ) )
				return current_section;
		}

		return nullptr;
	}
}