#pragma once
#include <string>
#include <memory>
#include "mem_defs.hpp"

namespace impl
{
	class unique_attachment
	{
	private:
		KAPC_STATE apc_state{};
		PEPROCESS process{};
	public:
		explicit unique_attachment( PEPROCESS process )
		{
			if ( !process )
				return;

			KeStackAttachProcess( process, &apc_state );
		}

		~unique_attachment( )
		{
			KeUnstackDetachProcess( &apc_state );
			ObfDereferenceObject( process );
		}
	};

	bool write_to_read_only( void* destination, void* source, std::size_t size )
	{
		const std::unique_ptr<MDL, decltype( &IoFreeMdl )> mdl( IoAllocateMdl( destination, static_cast< ULONG >( size ), FALSE, FALSE, nullptr ), &IoFreeMdl );

		if ( !mdl )
			return false;

		MmProbeAndLockPages( mdl.get( ), KernelMode, IoReadAccess );

		const auto mapped_page = MmMapLockedPagesSpecifyCache( mdl.get( ), KernelMode, MmNonCached, nullptr, FALSE, NormalPagePriority );

		if ( !mapped_page )
			return false;

		if ( !NT_SUCCESS( MmProtectMdlSystemAddress( mdl.get( ), PAGE_EXECUTE_READWRITE ) ) )
			return false;

		std::memcpy( mapped_page, source, size );

		MmUnmapLockedPages( mapped_page, mdl.get( ) );
		MmUnlockPages( mdl.get( ) );

		return true;
	}

	template <typename T = std::uint8_t*>
	__forceinline T follow_call( std::uint8_t* address )
	{
		/* + 1 is the address of the calle, + 5 is the size of a call instruction */
		return ( T )( address + *reinterpret_cast< std::int32_t* >( address + 1 ) + 5 );
	}

	template <typename T = std::uint8_t*>
	__forceinline T follow_conditional_jump( std::uint8_t* address )
	{
		/* + 1 is the offset of the jump, + 2 is the size of a conditional jump */
		return ( T )( address + *reinterpret_cast< std::int8_t* >( address + 1 ) + 2 );
	}	

	template <typename T = std::uint8_t*>
	__forceinline T resolve_mov( std::uint8_t* address )
	{
		/* + 3 is the address of the source, + 7 is the size of a mov instruction */
		return ( T )( address + *reinterpret_cast<std::int32_t*>( address + 3 ) + 7 );
	}
}