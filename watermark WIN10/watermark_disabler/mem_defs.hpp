#pragma once
#include <cstdint>
#include <ntifs.h>
#include <stdlib.h>

namespace nt
{
    struct rtl_module_info
    {
        char pad_0[ 0x10 ];
        PVOID image_base;
        ULONG image_size;
        char pad_1[ 0xa ];
        USHORT file_name_offset;
        UCHAR full_path[ _MAX_PATH - 4 ];
    };

    struct rtl_modules
    {
        ULONG count;
        rtl_module_info modules[ 1 ];
    };

	struct image_file_header
	{
		USHORT machine;
		USHORT number_of_sections;
	};

	struct image_section_header
	{
		std::uint8_t  name[ 8 ];

		union
		{
			std::uint32_t physical_address;
			std::uint32_t virtual_size;
		} misc;

		std::uint32_t virtual_address;
		std::uint32_t size_of_raw_data;
		std::uint32_t pointer_to_raw_data;
		std::uint32_t pointer_to_relocations;
		std::uint32_t pointer_to_line_numbers;
		std::uint16_t number_of_relocations;
		std::uint16_t number_of_line_numbers;
		std::uint32_t characteristics;
	};

	struct image_nt_headers
	{
		std::uint32_t signature;
		image_file_header file_header;
	};
}