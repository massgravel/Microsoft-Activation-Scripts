#pragma once
#include <memory>
#include <ntifs.h>

namespace impl
{
	struct unique_pool
	{
		void operator( )( void* pool )
		{
			if ( pool )
				ExFreePoolWithTag( pool, 0 );
		}
	};

	using pool = std::unique_ptr<void, unique_pool>;

	struct unique_object
	{
		void operator( )( void* object )
		{
			if ( object )
				ObfDereferenceObject( object );
		}
	};

	template <typename T>
	using object = std::unique_ptr<std::remove_pointer_t<T>, unique_object>;
}
