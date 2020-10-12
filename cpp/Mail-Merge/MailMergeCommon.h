#pragma once

#include <system/object_ext.h>


template<typename T>
System::SharedPtr<System::Array<System::SharedPtr<System::Object>>> BoxVector(const std::vector<T>& strings)
{
	auto result = System::MakeArray<System::SharedPtr<System::Object>>(strings.size());
	std::transform(strings.begin(), strings.end(), result.begin(),
		[](const T& value)
		{
			return System::ObjectExt::Box<T>(value);
		});

	return result;
}

template<typename ... Ts> System::SharedPtr<System::Array<System::SharedPtr<System::Object>>> Box(Ts ... args)
{
	return System::MakeArray<System::SharedPtr<System::Object>>({System::ObjectExt::Box(args)...});
}
