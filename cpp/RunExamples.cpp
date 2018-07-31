#include "stdafx.h"

#include "RunExamples.h"

#include <system/io/path.h>

using namespace System;
using namespace System::IO;

String GetDataDir_QuickStart()
{
    return Path::GetFullPath(u"Data/Quick-Start/");
}

String GetOutputFilePath(const String& inputFilePath)
{
    return  Path::GetFileNameWithoutExtension(inputFilePath) + u"_out" + Path::GetExtension(inputFilePath);
}

String GetDataDir_LINQ()
{
    return Path::GetFullPath(u"Data/LINQ/");
}

String GetDataDir_Database()
{
    return Path::GetFullPath(u"Data/Database/");
}

String GetDataDir_LoadingAndSaving()
{
    return Path::GetFullPath(u"Data/Loading-and-Saving/");
}

String GetDataDir_JoiningAndAppending()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Joining-Appending/");
}

String GetDataDir_WorkingWithList()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Lists/");
}

String GetDataDir_FindAndReplace()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Find-Replace/");
}

String GetDataDir_ConvertUtil()
{
    return Path::GetFullPath(u"Data/Programming-Documents/ConvertUtil/");
}

String GetDataDir_WorkingWithBookmarks()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Bookmarks/");
}

String GetDataDir_WorkingWithComments()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Comments/");
}

String GetDataDir_WorkingWithDocument()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Document/");
}

String GetDataDir_WorkingWithShapes()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Shapes/");
}

String GetDataDir_WorkingWithFields()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Fields/");
}

String GetDataDir_WorkingWithHyperlink()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Hyperlink/");
}

String GetDataDir_WorkingWithCharts()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Charts/");
}

String GetDataDir_WorkingWithOnlineVideo()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Video/");
}

String GetDataDir_WorkingWithNode()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Node/");
}

String GetDataDir_WorkingWithTheme()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Theme/");
}

String GetDataDir_WorkingWithRanges()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Ranges/");
}

String GetDataDir_WorkingWithStructuredDocumentTag()
{
    return Path::GetFullPath(u"Data/Programming-Documents/StructuredDocumentTag/");
}

String GetDataDir_WorkingWithImages()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Images/");
}

String GetDataDir_WorkingWithStyles()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Styles/");
}

String GetDataDir_WorkingWithTables()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Tables/");
}

String GetDataDir_WorkingWithSignature()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Signature/");
}

String GetDataDir_WorkingWithSections()
{
    return Path::GetFullPath(u"Data/Programming-Documents/Sections/");
}

String GetDataDir_MailMergeAndReporting()
{
    return Path::GetFullPath(u"Data/Mail-Merge/");
}


String GetDataDir_RenderingAndPrinting()
{
    return Path::GetFullPath(u"Data/Rendering-Printing/");
}

String GetDataDir_ViewersAndVisualizers()
{
    return Path::GetFullPath(u"Data/Viewers-Visualizers/");
}