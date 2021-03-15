#include "stdafx.h"
#include "examples.h"
#include <system/string.h>
#include <Aspose.Words.Cpp/Model/Document/Document.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaProject.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReferenceCollection.h>
#include <Aspose.Words.Cpp/RW/Ole/Vba/VbaReference.h>
#include <system/enumerator_adapter.h>

namespace
{

//ExStart: GetLibIdAndReferencePath
/// <summary>
/// Returns path from a specified identifier of an Automation type library.
/// </summary>
/// <remarks>
/// Please see details for the syntax at [MS-OVBA], 2.1.1.8 LibidReference. 
/// </remarks>
System::String GetLibIdReferencePath(const System::String& libIdReference)
{
	if (!System::String::IsNullOrEmpty(libIdReference))
	{
		auto refPaths = libIdReference.Split(u'#');
		if (refPaths->get_Length() > 3)
		{
			return refPaths[3];
		}
	}

	return u"";
}

/// <summary>
/// Returns path from a specified identifier of an Automation type library.
/// </summary>
/// <remarks>
/// Please see details for the syntax at [MS-OVBA], 2.1.1.12 ProjectReference. 
/// </remarks>
System::String GetLibIdProjectPath(const System::String& libIdProject)
{
	return !System::String::IsNullOrEmpty(libIdProject) ? libIdProject.Substring(3) : u"";
}

/// <summary>
/// Returns string representing LibId path of a specified reference. 
/// </summary>
System::String GetLibIdPath(const System::SharedPtr<Aspose::Words::Vba::VbaReference>& reference)
{
	switch (reference->get_Type())
	{
	case Aspose::Words::Vba::VbaReferenceType::Registered: 
	case Aspose::Words::Vba::VbaReferenceType::Original: 
	case Aspose::Words::Vba::VbaReferenceType::Control: 
		return GetLibIdReferencePath(reference->get_LibId());
	case Aspose::Words::Vba::VbaReferenceType::Project: 
		return GetLibIdProjectPath(reference->get_LibId());
	default: ;
		throw System::ArgumentOutOfRangeException();
	}
}
//ExEnd: GetLibIdAndReferencePath


}

void WorkingWithVbaReferenceCollection()
{
	// The path to the documents directory.
	auto const inputDataDir = GetInputDataDir_LoadingAndSaving();
	
	//ExStart: RemoveReferenceFromCollectionOfReferences
	auto doc = System::MakeObject<Aspose::Words::Document>(inputDataDir + u"VbaProject.docm");

	
	// Find and remove the reference with some LibId path.
	const System::String brokenPath = u"brokenPath.dll";
	auto references = doc->get_VbaProject()->get_References();
	
	// Find and remove the reference with some LibId path.
	for (auto reference : System::IterateOver(doc->get_VbaProject()->get_References()))
	{
		auto path = GetLibIdPath(reference);
		if (path == brokenPath)
		{
			references->Remove(reference);
		}
	}

	doc->Save(GetOutputDataDir_LoadingAndSaving() + u"NoBrokenRef.docm");
	//ExEnd: RemoveReferenceFromCollectionOfReferences
}
