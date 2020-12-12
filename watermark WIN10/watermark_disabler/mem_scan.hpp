#pragma once
#include "mem_util.hpp"
#include "mem_iter.hpp"
#include "util_raii.hpp"

namespace impl
{
	inline bool search_for_signature_helper( const std::uint8_t* data, const std::uint8_t* signature, const char* mask )
	{
		// check if page is correct & readable (this internally checks PTE, PDE, ...)
		if ( !MmIsAddressValid( const_cast< std::uint8_t* >( data ) ) )
			return false;

		// iterate through validity of the mask (mask & signature are equal
		for ( ; *mask; ++mask, ++data, ++signature )
				if ( *mask == 'x' && *data != *signature ) // if mask is 'x' (a match), and the current byte is not equal to the byte in the signature, then return false.
					return false;

		return true;
	}
	
	std::uint8_t* search_for_signature( const nt::rtl_module_info* module, const char* signature, const char* signature_mask )
	{
		if ( !module )
			return nullptr;

		const auto module_start = reinterpret_cast< std::uint8_t* >( module->image_base );
		const auto module_size = module_start + module->image_size;

		/* iterate the entire module */
		for ( auto segment = module_start; segment < module_size; segment++ )
		{
			if ( search_for_signature_helper( segment, reinterpret_cast< std::uint8_t* >( const_cast< char* >( signature ) ), signature_mask ) )
				return segment;
		}

		return nullptr;
	}

	extern "C" NTSYSAPI PCHAR NTAPI PsGetProcessImageFileName( PEPROCESS );

	PEPROCESS search_for_process( const char* process_name )
	{
		const auto kernel_module_info = search_for_module( "ntoskrnl.exe" );

		if ( !kernel_module_info )
			return nullptr;

		/* we are scanning for a conditional jump, that jumps to a call to the unexported function that we want, so we follow the jump, then follow the call to get to the function. */
		const auto conditional_instruction = search_for_signature( kernel_module_info, "\x79\xdc\xe9", "xxx" );

		if ( !conditional_instruction )
			return nullptr;
		
		const auto call_instruction = follow_conditional_jump( conditional_instruction );

		if ( !call_instruction )
			return nullptr;

		const auto PsGetNextProcess = follow_call< PEPROCESS( __stdcall* )( PEPROCESS ) >( call_instruction );

		if ( !PsGetNextProcess )
			return nullptr;

		PEPROCESS previous_process = PsGetNextProcess( nullptr );

		while ( previous_process )
		{
			if ( !std::strcmp( PsGetProcessImageFileName( previous_process ), process_name ) )
				return previous_process;

			previous_process = PsGetNextProcess( previous_process );
		}

		return nullptr;
	}
}
