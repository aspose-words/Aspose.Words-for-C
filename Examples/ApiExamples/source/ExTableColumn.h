#pragma once
// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#include <system/collections/list.h>
#include <system/array.h>
#include <cstdint>
#include <Aspose.Words.Cpp/Model/Tables/Table.h>
#include <Aspose.Words.Cpp/Model/Tables/Cell.h>

#include "ApiExampleBase.h"


using namespace Aspose::Words::Tables;

namespace Aspose {

namespace Words {

namespace ApiExamples {

class ExTableColumn : public ApiExampleBase
{
    typedef ExTableColumn ThisType;
    typedef ApiExampleBase BaseType;
    
    typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
    RTTI_INFO_DECL();
    
public:

    /// <summary>
    /// Represents a facade object for a column of a table in a Microsoft Word document.
    /// </summary>
    class Column : public System::Object
    {
        typedef Column ThisType;
        typedef System::Object BaseType;
        
        typedef ::System::BaseTypesInfo<BaseType> ThisTypeBaseTypesInfo;
        RTTI_INFO_DECL();
        
    public:
    
        /// <summary>
        /// Returns the cells which make up the column.
        /// </summary>
        System::ArrayPtr<System::SharedPtr<Aspose::Words::Tables::Cell>> get_Cells();
        
        /// <summary>
        /// Returns a new column facade from the table and supplied zero-based index.
        /// </summary>
        static System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> FromIndex(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex);
        /// <summary>
        /// Returns the index of the given cell in the column.
        /// </summary>
        int32_t IndexOf(System::SharedPtr<Aspose::Words::Tables::Cell> cell);
        /// <summary>
        /// Inserts a new column before this column into the table.
        /// </summary>
        System::SharedPtr<Aspose::Words::ApiExamples::ExTableColumn::Column> InsertColumnBefore();
        /// <summary>
        /// Removes the column from the table.
        /// </summary>
        void Remove();
        /// <summary>
        /// Returns the text of the column. 
        /// </summary>
        System::String ToTxt();
        
    private:
    
        int32_t mColumnIndex;
        System::SharedPtr<Aspose::Words::Tables::Table> mTable;
        
        Column(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex);
        
        MEMBER_FUNCTION_MAKE_OBJECT_DECLARATION(Column, CODEPORTING_ARGS(System::SharedPtr<Aspose::Words::Tables::Table> table, int32_t columnIndex));
        /// <summary>
        /// Provides an up-to-date collection of cells which make up the column represented by this facade.
        /// </summary>
        System::SharedPtr<System::Collections::Generic::List<System::SharedPtr<Aspose::Words::Tables::Cell>>> GetColumnCells();
        
    };
    
    
public:

    void RemoveColumnFromTable();
    void Insert();
    void TableColumnToTxt();
    
};

} // namespace ApiExamples
} // namespace Words
} // namespace Aspose


