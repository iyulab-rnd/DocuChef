Basic Concepts
==============

### Templates

All work on creating a report is built on templates (XLSX-templates) - Excel books that contain a description of the report form, as well as options for books, sheets and report areas. Special field formulas and data areas describe the data in the report structure that you want to transfer to Excel. ClosedXML.Report will call the template and fill the report cells with data from the specified sets.

### Variables

The values passed in ClosedXML.Report with the method `AddVariable` are called variables. They are used to calculate the expressions in the templates. Variable can be added with or without a name. If a variable is added without a name, then all its public fields and properties are added as variables with their names.

Examples:

`template.AddVariable(cust);`

OR

`template.AddVariable("Customer", cust);`.

### Expressions

Expressions are enclosed in double braces {{ }} and utilize the syntax similar to C#. Lambda expressions are supported.

Examples:

`{{item.Product.Price * item.Product.Quantity}}`

`{{items.Where(i => i.Currency == "RUB").Count()}}`

### Tags

ClosedXML.Report has a few advanced features allowing to hide a worksheet, sort the data table, apply groupping, calculate totals, etc. These features are controlled by addit tags to the worksheet, to the entire range, or to a single column. Tag is a text embrased by double angle brackets that can be analyzed by a ClosedXML.Report parser. Different tags let you get subtotals, build pivot tables, apply auto-filter, and so on. Tags may have parameters for tuning their behavior. Parameters may require you to specify their values. In this case, the parameter name is separated from the value by the equal sign.

All tags can refer to six report objects: report, sheet, column, row, region, column of area. Report tags are specified in cell A1 of any template sheet. Sheet tags are specified in cell A2 on the sheet. The column’s tags are specified in the first line of the sheet. The line tags are specified in the first column of the sheet. The area tags are specified in the leftmost cell of the options row of this area. Tags of the column of the area are specified in the cell of this column in the options row of the area.

Example: `<<Range horizontal>>`.

A list of all the tags you can find on the [Tags page](More-options)

### Ranges

To represent IEnumerable values, Excel named regions are used.

Quick Start
===========

ClosedXML.Report is a tool for report generation and data analysis in .NET applications through the use of Microsoft Excel.

It is a .NET-library for report generation Microsoft Excel without requiring Excel to be installed on the machine that’s running the code. With ClosedXML.Report, you can easily export any data from your .NET classes to Excel using the XLSX-template.

Excel is an excellent alternative to common report generators, and using Excel’s built-in features can make your reports much more responsive. Use ClosedXML.Report as a tool for generating Excel files. Then use Excel visual instruments: formatting (including conditional formatting), AutoFilter, and Pivot tables to construct a versatile data analysis system. With ClosedXML.Report, you can move a lot of report programming and tuning into Excel. ClosedXML.Report templates are simple and our algorithms are fast – we carefully count every millisecond – so you waste less time on routine report programming and get surprisingly fast results. If you want to master such a versatile tool as Excel – ClosedXML.Report is an excellent choice. Furthermore, ClosedXML.Report doesn’t operate with the usual concepts of band-oriented report tools: Footer, Header, and Detail. So you get a much greater degree of freedom in report construction and design, and the easiest possible integration of .NET and Microsoft Excel.

### Install ClosedXML.Report via NuGet

If you want to include ClosedXML.Report in your project, you can [install it directly from NuGet](https://www.nuget.org/packages/ClosedXML.Report/)

To install ClosedXML.Report, run the following command in the Package Manager Console

    PM> Install-Package ClosedXML.Report
    

or if you have a signed assembly, then use:

    PM> Install-Package ClosedXML.Report.Signed
    

Features
--------

*   Copying cell formatting
*   Propagation of conditional formatting
*   Vertical and horizontal tables and subranges
*   Ability to implement Excel formulas
*   Using dynamically calculated formulas with the syntax of C # and Linq
*   Operations with tabular data: sorting, grouping, total functions.
*   Pivot tables
*   Subranges

How to use?
-----------

To create a report you must first create a report template. You can apply any formatting to any workbook cells, insert pictures, and modify any of the parameters of the workbook itself. In this example, we have turned off the zero values display and hidden the gridlines. ClosedXML.Report will preserve all changes to the template.

**Template**

![template1](/ClosedXML.Report/images/quick-start-01.png)

**Code**

        protected void Report()
        {
            const string outputFile = @".\Output\report.xlsx";
            var template = new XLTemplate(@".\Templates\report.xlsx");
    
            using (var db = new DbDemos())
            {
                var cust = db.customers.LoadWith(c => c.Orders).First();
                template.AddVariable(cust);
                template.Generate();
            }
    
            template.SaveAs(outputFile);
    
            //Show report
            Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });
        }
    

**Result**

![result1](/ClosedXML.Report/images/quick-start-02.png)

For more information see [the documentation](index) and [tests](https://github.com/ClosedXML/ClosedXML.Report/tree/master/tests)

Simple Template
===============

You can use _expressions_ with braces {{ }} in any cell of any sheet of the _template_ workbook and Excel will find their values at run-time. How? ClosedXML.Report adds a hidden worksheet in a report workbook and transfers values of all fields for the current record. Then it names all these data cells.

Excel formulas, in which _variables_ are added, must be escaped `&`. As an example: `&=CONCATENATE({{Addr1}}, " ", {{Addr2}})`. Pay attention that escaped formulas must follow international, non-localized syntax which includes English names of functions, comma as a value separator, etc. For example, in the Russian version of Excel you could use formula `=СУММ(A1:B2; E1:F2)` for ordinary cells, but the escaped formula should be `&=SUM(A1:B2, E1:F2)`.

Cells with field formulas can be formatted in any known way, including conditional formatting.

Here is a simple example:

![simpletemplate](/ClosedXML.Report/images/simple-template-01.png)

    ...
            var template = new XLTemplate(@".\Templates\template.xlsx");
            var cust = db.Customers.GetById(10);
    
            template.AddVariable(cust);
            // OR
            template.AddVariable("Company", cust.Company);
            template.AddVariable("Addr1", cust.Addr1);
            template.AddVariable("Addr2", cust.Addr2);
    ...
    
    public class Customer
    {
    	public double CustNo { get; set; }
    	public string Company { get; set; }
    	public string Addr1 { get; set; }
    	public string Addr2 { get; set; }
    	public string City { get; set; }
    	public string State { get; set; }
    	public string Zip { get; set; }
    	public string Country { get; set; }
    	public string Phone { get; set; }
    	public string Fax { get; set; }
    	public double? TaxRate { get; set; }
    	public string Contact { get; set; }
    }

Report examples
===============

### Simple Template

![simple](/ClosedXML.Report/images/examples-01.png)

You can apply to cells any formatting including conditional formats.

The template: [Simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Simple.xlsx)

The result file: [Simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/Simple.xlsx)

### Sorting the Collection

![tlists1_sort](/ClosedXML.Report/images/examples-02.png)

You can sort the collection by columns. Specify the tag `<<sort>>` in the options row of the corresponding columns. Add option `desc` to the tag if you wish the list to be sorted in the descending order (`<<sort desc>>`).

For more details look to the [Sorting](Sorting)

The template: [tLists1\_sort.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tLists1_sort.xlsx)

The result file: [tLists1\_sort.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tLists1_sort.xlsx)

### Totals

![tlists2_sum](/ClosedXML.Report/images/examples-03.png)

You can get the totals for the column in the ranges by specifying the tag in the options row of the corresponding column. In the example above we used tag `<<sum>>` in the column Amount paid.

For more details look to the [Totals in a Column](Totals-in-a-column).

The template: [tlists2\_sum.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tLists2_sum.xlsx)

The result file: [tlists2\_sum.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tLists2_sum.xlsx)

### Range and Column Options

![tLists3_options](/ClosedXML.Report/images/examples-04.png)

Besides specifying the data for the range ClosedXML.Report allows you to sort the data in the range, calculate totals, group values, etc. ClosedXML.Report performs these actions if it founds the range or column tags in the service row of the range.

For more details look to the [Flat Tables](Flat-tables)

In the example above example we applied auto filters, specified that columns must be resized to fit contents, replaced Excel formulas with the static text and protected the “Amount paid” against the modification. For this, we used tags `<<AutoFilter>>`, `<<ColsFit>>`, `<<OnlyValues>>` and `<<Protected>>`.

The template: [tLists3\_options.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tLists3_options.xlsx)

The result file: [tLists3\_options.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tLists3_options.xlsx)

### Complex Range

![tlists4_complexrange](/ClosedXML.Report/images/examples-05.png)

ClosedXML.Report can use multi-row templates for the table rows. You may apply any format you wish to the cells, merge them, use conditional formats, Excel formulas.

For more details look to the [Flat Tables](Flat-tables)

The template: [tLists4\_complexRange.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tLists4_complexRange.xlsx)

The result file: [tLists4\_complexRange.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tLists4_complexRange.xlsx)

### Grouping

![GroupTagTests_Simple](/ClosedXML.Report/images/examples-06.png)

The `<<group>>` tag may be used along with any of the aggregating tags. Put the tag `<<group>>` into the service row of those columns which you wish to use for aggregation.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_Simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_Simple.xlsx)

The result file: [GroupTagTests\_Simple.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_Simple.xlsx)

### Collapsed Groups

![GroupTagTests_Collapse](/ClosedXML.Report/images/examples-07.png)

Use the parameter collapse of the group tag (`<<group collapse>>`) if you want to display only those rows that contain totals or captions of data sections.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_Collapse.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_Collapse.xlsx)

The result file: [GroupTagTests\_Collapse.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_Collapse.xlsx)

### Summary Above the Data

![GroupTagTests_SummaryAbove](/ClosedXML.Report/images/examples-08.png)

ClosedXML.Report implements the tag `summaryabove` that put the summary row above the grouped rows.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_SummaryAbove.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_SummaryAbove.xlsx)

The result file: [GroupTagTests\_SummaryAbove.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_SummaryAbove.xlsx)

### Merged Cells in Groups (option 1)

![GroupTagTests_MergeLabels](/ClosedXML.Report/images/examples-09.png)

The `<<group>>` tag has options making it possible merge cells in the grouped column. To achieve this specify the parameter mergelabels in the group tag (`<<group mergelabels>>`).

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_MergeLabels.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_MergeLabels.xlsx)

The result file: [GroupTagTests\_MergeLabels.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_MergeLabels.xlsx)

### Merged Cells in Groups (option 2)

![GroupTagTests_MergeLabels2](/ClosedXML.Report/images/examples-10.png)

Tag `<<group>>` allows to group cells without adding the group title. This function may be enabled by using parameter MergeLabels=Merge2 in the group tag (`<<group MergeLabels=Merge2>>`). Cells containing the grouped data are merged and filled with the group caption.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_MergeLabels2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_MergeLabels2.xlsx)

The result file: [GroupTagTests\_MergeLabels2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_MergeLabels2.xlsx)

### Nested Groups

![GroupTagTests_NestedGroups](/ClosedXML.Report/images/examples-11.png)

Ranges may be nested with no limitation on the depth of nesting.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_NestedGroups.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_NestedGroups.xlsx)

The result file: [GroupTagTests\_NestedGroups.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_NestedGroups.xlsx)

### Disable Groups Collapsing

![GroupTagTests_DisableOutline](/ClosedXML.Report/images/examples-12.png)

Use the option disableoutline of the group tag (`<<group disableoutline>>`) to prevent them from collapsing. In the example above the range is grouped by both Company and Payment method columns. Collapsing of groups for the Payment method column is disabled.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_DisableOutline.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_DisableOutline.xlsx)

The result file: [GroupTagTests\_DisableOutline.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_DisableOutline.xlsx)

### Specifying the Location of Group Captions

![GroupTagTests_PlaceToColumn](/ClosedXML.Report/images/examples-13.png)

The `<<group>>` tag has a possibility to put the group caption in any column of the grouped range by using the parameter `PLACETOCOLUMN=n` where `n` defines the column number in the range. (starting from 1). Besides, ClosedXML.Report supports the `<<delete>>` tag that aims to specify columns to delete. In the example above the Company column is grouped with the option `mergelabels`. The group caption is placed to the second column (`PLACETOCOLUMN=2`). Finally, the Company column is removed.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_PlaceToColumn.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_PlaceToColumn.xlsx)

The result file: [GroupTagTests\_PlaceToColumn.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_PlaceToColumn.xlsx)

### Formulas in Group Line

![GroupTagTests_FormulasInGroupRow](/ClosedXML.Report/images/examples-14.png)

ClosedXML.Report saves the full text of cells in the service row, except tags. You can use this feature to specify Excel formulas in group captions. In the example above there is grouping by columns Company and Payment method. The Amount Paid column contains an Excel formula in the service row.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_FormulasInGroupRow.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_FormulasInGroupRow.xlsx)

The result file: [GroupTagTests\_FormulasInGroupRow.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_FormulasInGroupRow.xlsx)

### Groups with Captions

![GroupTagTests_WithHeader](/ClosedXML.Report/images/examples-15.png)

You can configure the appearance of the group caption by using the `WITHHEADER` parameter of the `<<group>>` tag. With this, the group caption is placed over the grouped rows. The `SUMMARYABOVE` does not change this behavior.

For more details look to the [Grouping](Grouping)

The template: [GroupTagTests\_WithHeader.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/GroupTagTests_WithHeader.xlsx)

The result file: [GroupTagTests\_WithHeader.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/GroupTagTests_WithHeader.xlsx)

### Nested Ranges

![Subranges_Simple_tMD1](/ClosedXML.Report/images/examples-16.png)

You can place one ranges inside the others in order to reflect the parent-child relation between entities. In the example above the `Items` range is nested into the `Orders` range which, in turn, is nested to the `Customers` range. Each of three ranges has its own header, and all have the same left and right boundary.

For more details look to the [Nested ranges: Master-detail reports](Nested-ranges_-Master-detail-reports).

The template: [Subranges\_Simple\_tMD1.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Subranges_Simple_tMD1.xlsx)

The result file: [Subranges\_Simple\_tMD1.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/Subranges_Simple_tMD1.xlsx)

### Nested Ranges with Subtotals

![Subranges_WithSubtotals_tMD2](/ClosedXML.Report/images/examples-17.png)

You may use aggregation tags at any level of your master-detail report. In the example above the `<<sum>>` tag in the I9 cell will summarize columns \`\` in the scope of an order, while the same tag in the I10 cell will summarize all the data for each Customer.

For more details look to the [Nested ranges: Master-detail reports](Nested-ranges_-Master-detail-reports).

The template: [Subranges\_WithSubtotals\_tMD2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Subranges_WithSubtotals_tMD2.xlsx)

The result file: [Subranges\_WithSubtotals\_tMD2.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/Subranges_WithSubtotals_tMD2.xlsx)

### Nested Ranges with Sorting

![Subranges_WithSort_tMD3](/ClosedXML.Report/images/examples-18.png)

You can use the `<<sort>>` for the nested ranges as well.

For more details look to the [Nested ranges: Master-detail reports](Nested-ranges_-Master-detail-reports).

The template: [Subranges\_WithSort\_tMD3.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Subranges_WithSort_tMD3.xlsx)

The result file: [Subranges\_WithSort\_tMD3.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/Subranges_WithSort_tMD3.xlsx)

### Pivot Tables

![tPivot5_Static](/ClosedXML.Report/images/examples-19.png)

ClosedXML.Report support such a powerful tool for data analysis as pivot tables. You can define one or many pivot tables directly in the report template to benefit the power of the Excel pivot table constructor and nearly all the available features for they configuring and designing.

For more details look to the [Pivot Tables](Pivot-tables).

The template: [tPivot5\_Static.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tPivot5_Static.xlsx)

The result file: [tPivot5\_Static.xlsx](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Gauges/tPivot5_Static.xlsx)

Flat Tables
===========

Variables or their properties of type `IEnumerable` may be bounded to regions (_flat tables_). To output all values from `IEnumerable` you should create a named range with the same name as the variable. ClosedXML.Report searches named ranges and maps variables to them. To establish binding to properties of collection elements use the built-in name `item`.

There are certain limitations on range configuration:

*   Ranges can only be rectangular
*   Ranges must not have gaps
*   Cells in ranges may store normal text, ClosedXML.Report _expressions_ (in double curly braces), standard Excel formulas, and formulas escaped with `&` character
*   Cells in ranges may be empty

### Range names

While building a document, ClosedXML.Report finds all named ranges and determines data sources by their name. Range name should coincide with the name of the variable serving a data source for this range. For nested tables, the range name is built using an underscore (`_`). E.g., to output values from `Customers[].Orders[].Items[]` the range name must be `Customers_Orders_Items`. This example may be found in the [sample template](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Subranges_Simple_tMD1.xlsx).

### Expressions within tables

To work with tabular data, ClosedXML.Report introduces special variables that you can use in expressions inside tables:

*   `item` - element from the list.
*   `index` - the index of the element in the list (starts with 0).
*   `items` - the entire list of items.

Vertical tables
---------------

**Requirements for vertical tables**

Each range specifying a vertical table must have at least two columns and two rows. The leftmost column and bottommost row serve to hold configuration information and are treated specially. After the report is built the service column is cleared, and the service row is deleted if it is empty.

When dealing with vertical tables, CLosedXML.Report performs the following actions:

*   The required number of rows is inserted in the region. Note that cells are added to the range only, not to the entire worksheet. This means that regions located to the right or left of the table won’t be affected.
*   Contents of the added cells are filled with data according to the template (static text, formulas or _expressions_).
*   Styles from the template are applied to the inserted cells.
*   Template cells are deleted.
*   If the service row does not contain any options then it is deleted too.
*   If there are options defined in the service row then they are processed accordingly, and then the row is either cleared or deleted.

Look at the example from the [start page](https://github.com/ClosedXML/ClosedXML.Report). As you can see in the picture, there is a range named `Orders` including a single row with _expressions_, a service row and a service column.

![template](/ClosedXML.Report/images/flat-tables-01.png)

We applied custom styles to the cells in the range, i.e. we specified date formats for cells `SaleDate` and `ShipDate` and number formats with separators for cells `Items total` и `Amount paid`. In addition, we applied a conditional format to the `Payment method` cell.

To build a report from the template you simply have to run this code:

    ...
            var template = new XLTemplate('template.xslx');
            var cust = db.Customers.GetById(10);
    
            template.AddVariable(cust);
            template.Generate();
    ...
    
    public class Customer
    {
        ...
        public List<order> Orders { get; set; }
    }
    
    public class order
    {
    	public int OrderNo { get; set; } 
    	public DateTime? SaleDate { get; set; } // DateTime
    	public DateTime? ShipDate { get; set; } // DateTime
    	public string ShipToAddr1 { get; set; } // text(30)
    	public string ShipToAddr2 { get; set; } // text(30)
    	public string PaymentMethod { get; set; } // text(7)
    	public double? ItemsTotal { get; set; } // Double
    	public double? TaxRate { get; set; } // Double
    	public double? AmountPaid { get; set; } // Double
    }
    

In the picture below you see the report produced from the specified template. Note that selected area now contains the data and is named `Orders`. You can use this name to access data in the result report.

![result](/ClosedXML.Report/images/flat-tables-02.png)

Horizontal tables
-----------------

Horizontal tables do not have such strict requirements as vertical tables. The named range consists of a single row (in other words, without an options row) and is assumed to be a horizontal table. A horizontal table does not need to have a service column either. In fact, the horizontal table may be defined by a single cell. To explicitly define a range as a horizontal table definition, put a special tag `<<Range horizontal>>` into any cell inside the range. You may find the example using the horizontal range [on the GitHub](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/4.xlsx).

There two ranges in that template - `dates` and `PlanData_Hours`. Each of these ranges consist of one cell. As has been said, ClosedXML.Report treats such ranges as horizontal table definitions.

![horizontal template](/ClosedXML.Report/images/flat-tables-03.png)

The result report:

![horizontal result](/ClosedXML.Report/images/flat-tables-04.png)

Service row
-----------

ClosedXML.Report offers nice features for data post-processing on building report: it sort the data, calculate totals by columns, apply grouping, etc. Which actions to perform may be defined by putting special _tags_ to the template. _Tag_ is a keyword put into double angle brackets along with configuration options. Tags controlling data in tables should be placed in the service row of the table template. Some of the tags are aplied to the range as a whole, the others affect only a column they are put in.

We will give detailed information on tags usage in the next chapters.

Now consider the following template.

![simple template](/ClosedXML.Report/images/flat-tables-05.png)

The cell in the service row in the `Amount paid` column contains the tag `<<sum>>`. The cell next to it contains a static text “Total”. After the report is built the tag `<<sum>>` will be replaced with a formula calculating the sum of the amounts of the entire column. Tag `<<sum>>` belongs to the “column” tags. Such tags are applied to the column they are put in. Other examples of the “column” tags are `<<sort>>` that defines the ordering of the data set by the specified column, or `<<group>>` configuring grouping by the specified field.

Actions that must be performed on the whole range are defined with “range” tags. They are defined in the first (leftmost) cell of the service row. You may experiment a little with the described template. Try to open it and write tags `<<Autofilter>> <<OnlyValues>>` into the first cell of the service row. After you saved the template and rebuilt the report you may see that now it has the auto-filter turned on, and the formula `=SUBTOTAL(9, ...` in the `Amount paid` column has been replaced with the static value.

Multiple ranges support for Named Ranges
----------------------------------------

If we need to use one data source for several tables, then we can create a composite named range. In the example below, the Orders range includes two ranges `$A$5:$I$6` и `$A$10:$I$11`.

![image](https://user-images.githubusercontent.com/1150085/84245559-e6017080-ab0d-11ea-9643-b6fc5f3fd3e9.png)

When building a report, ClosedXML.Report will fill both of these ranges with data from the Orders variable.

Pivot Tables
============

To build pivot tables, it is sufficient to specify pivot table tags in the data range. After that, this range becomes the data source for the pivot table. The `<<pivot>>` tag is the first tag ClosedXML.Report pays attention to when analyzing cells in a data region. This tag can have multiple arguments. Here is the syntax:

`<<pivot Name=PivotTableName [Dst=Destination] [RowGrand] [ColumnGrand] [NoPreserveFormatting] [CaptionNoFormatting] [MergeLabels] [ShowButtons] [TreeLayout] [AutofitColumns] [NoSort]>>`

*   Name=PivotTableName is the name of the pivot table allowed in Excel.
*   Dst=Destination - the cell in which you want to place the left upper corner of the pivot table. If the Destination is not specified, then the pivot table is automatically placed on a new sheet of the book.
*   RowGrand - allows you to include in the pivot table the totals by rows.
*   ColumnGrand - includes totals for the pivot table.
*   NoPreserveFormatting - allows you to build a pivot table without preserving the formatting of the source range, which reduces the time to build the report.
*   CaptionNoFormatting - Formats the pivot table header in accordance with the source table.
*   MergeLabels - allows you to merge cells.
*   ShowButtons - shows a button to collapse and expand lines.
*   TreeLayout - sets the mode of the pivot table as a tree.
*   AutofitColumns - enables automatic selection of the width of the pivot table columns.
*   NoSort - disables automatic sorting of the pivot table.

Here are some examples of the correct setting of the Pivot option:

*   `<<pivot Name=Pivot1 Dst=Totals!A1>>` – a pivot table will be created with the name Pivot1; table will be placed on the Totals sheet starting at cell A1;
*   `<<pivot Name=Pivot25>>` – a pivot table will be created with the name Pivot25;
*   `<<pivot Name=Pivot25 Dst=Totals!A1 RowGrand>>` – the Pivot25 pivot table includes the totals for data lines;
*   `<<pivot Name=Pivot25 ColumnGrand>>` – the pivot table will include the totals for the columns.

Fields in all ranges of the pivot table are added in the order in which they appear in the template (from left to right). Therefore, when designing a data range on which a pivot table will be built, you need to adhere to one simple rules: line up the columns in the order in which you would like to see them in the pivot table

### Important!

The names of the fields for the pivot table are taken from the line above the data range - the heading of the source table. Be careful when creating this header, as there are some restrictions on the naming of fields in the pivot tables. With the help of pivot tables, it’s easy to create the most complicated cross-tables in reports.

### Template example

![template](/ClosedXML.Report/images/pivot-tables-01.png)

[Template file](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tPivot1.xlsx)

In the lower left cell of the data range there is a tag `<<pivot Name="OrdersPivot" dst="Pivot!B8" rowgrand mergelabels AutofitColumns>>`. This option will indicate ClosedXML.Report that a pivot table with the name “OrdersPivot” will be built across the region, which will be placed on the “Pivot” sheet starting at cell B8. And the parameter `rowgrand` will allow to include the totals for the columns of the resulting pivot table. In the service cell of the columns “Payment method”, “OrderNo”, “Ship date” and “Tax rate” the tag is `<<row>>`. The `<<row>>` tag defines the fields of the pivot table row area. In order to get the totals grouped by the method of payment of bills, the tag `<<sum>>` has been added to the tag `<<row>>` in the field “Payment method”. For the “Amount paid” and “Items total” fields, the `<<data>>` tag is specified (fields of the pivot table data range). In the options of the “Company” field, a `<<page>>` tag has been added (the page area field). When designing a template, in addition to the allocation of tags between the columns, do not forget to specify different formats for the cells of the range (including for cells with dates and numbers). Moreover, we formatted the service cells with column options, meaning that it is with this format that we will get subtotals in the pivot table. And for the “Payment method” field, we selected a cell with tags in color.

Static Pivot Tables
-------------------

You can place one or several pivot tables right in the report template, taking advantage of the convenience of the Excel Pivot Table wizard and virtually all the possibilities in their design and structuring. Let’s give an example. As a starting point, we use the [first example template](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/tPivot1.xlsx) with a summary table with the original Orders range on the Sheet1 sheet. Right in the template, we placed a static pivot table built over this range. The following figures show the steps for building this table. First, you need to select the source range for the pivot table. It is not identical to the Orders range, since it includes only the data line and the title above it. Notice how the source range is highlighted in the figure:

![pivot range](/ClosedXML.Report/images/pivot-tables-02.png)

Next, we put the pivot table on a separate PivotSheet and distributed its fields in the rows, columns, and data ranges. We formatted pivot table fields, as well as their headings. Finally, we called the pivot table as PivotTable1, and as an option to the source range, we specified `<<pivot>>`. After the data is transferred, all summary tables referencing this data range will be updated. That is, for one range you can build several pivot tables.

Nested Ranges: Master-detail Report
===================================

ClosedXML.Report makes possible creating reports with a comprehensive structure. This can be achieved by placing one range inside the other which corresponds to “one-to-many” relation between two sets of data.

Consider the example. We want to desing a report in such a way that table `Items` were nested into table `Orders`, which in its turn would be nested in `Customer` table. Given that, and because the main information belongs to the table `Items` let’s get started from it. On the picture you see the table `Items` definition.

![step 1](/ClosedXML.Report/images/nested-ranges-01.png)

As you can see, we named the range Customers\_Orders\_Items. Rules for assigning names to the ranges were explained in the section “[Flat tables#Range names](Flat-tables#Range-names)”. You also can see that we put a title above the table and left some place to the left of it, but that’s just the matter of appearance, you don’t have to do it if it’s not what need.

Then, the range `Customers_Orders_Items` must be placed inside the range `Customers_Orders`.

![step 2](/ClosedXML.Report/images/nested-ranges-02.png)

We inserted an empty string before the range, put a couple of _expressions_ here, and defined a new named range covering entire range `Customers_Orders_Items`, plus the row with captions for this table, plus the row below the table, assuming this row to be a service row for our new range. We also added another row to hold captions for `Orders` table (“Order no”, “Sale date”, etc.)

Finally, we started to build `Customers` table.

![step 3](/ClosedXML.Report/images/nested-ranges-03.png)

This time we did the very same steps: insert the row, define the _expressions_ (`{{CustNo}}` and `{{Company}}`), apply styles we like and define a named range `Customers` covering the nested ranges, plus a service row.

You can download the template file from [the GitHub](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Subranges_Simple_tMD1.xlsx)

To conclude, the rules to follow when creating a report with nested ranges are these:

*   All nested areas must be continuous.
*   Each range must have its own service row.
*   All ranges must have the same left and right boundaries; thus their widths must be equal too.
*   The leftmost column of all ranges is a “service” one.
*   One range may have any number of nested ranges.
*   The maximum depth of nested ranges is only limited by a worksheet capacity, when the data is specified.
*   The nested range may be places between any two rows of the outer range.
*   Each range in the hierarchy is treated as a whole, excluding the nested ranges, as described in the previous chapters.

Subtotals in Master-Detail Reports
----------------------------------

ClosedXML.Report supports calculating subtotals for columns in a range. Here we use tag `<<sum>>` in both ranges `Customers_Orders_Items` and `Customers_Orders`.

![totals](/ClosedXML.Report/images/nested-ranges-04.png)

The resulting template may be found on [the GitHub](https://github.com/ClosedXML/ClosedXML.Report/blob/develop/tests/Templates/Subranges_WithSubtotals_tMD2.xlsx)


Sorting a list range
====================

Regions in ClosedXML.Report may be sorted by columns. This can be done by using tag `<<sort>>` in the target columns of the service row.

![tlists1_sort](/ClosedXML.Report/images/sorting-01.png)

You may choose a descending order by adding the parameter `desc` to the tag (`<<sort desc>>`). When the report is built you’ll see the data sorted in descending order.

![tlists1_sort_desc](/ClosedXML.Report/images/sorting-02.png)

You also may add additional columns to the sort by adding the parameter `num` to the tag `sort` (`<<sort num=2>>`). On the picture below you may see the dataset will be sorted first by the column Payment Method, and then by the column Ship Date in the descending order.

![tlists1_sort_num xlsx](/ClosedXML.Report/images/sorting-03.png)

Sorting in ClosedXML.Report has this limitation: all the data of the range to order must be located in a single row. If this is not a option for you consider sorting your data before transferring it to ClosedXML.Report.

Totals in a Column
==================

In order to get the totals for the column in ClosedXML.Report there are aggregation tags:

*   SUM - displays the amount of the column;
*   AVG or AVERAGE - the average value of the column;
*   COUNT - the number of values in a column;
*   COUNTNUMS - the number of non-empty values in the column;
*   MAX - the maximum value in the column;
*   MIN - the minimum value in the column;
*   PRODUCT - product by column;
*   STDEV - standard deviation;
*   STDEVP - standard deviation of the total population
*   VAR - dispersion;
*   VARP - dispersion for the general population.

To calculate the results of ClosedXML.Report uses Excel tools, i.e. Each of these tags will be replaced by the corresponding Excel formula. For example, to calculate the Amount paid amount, we need to add the `<<sum>>` tag to the options line.

![tlists1_sum](https://user-images.githubusercontent.com/1150085/41203072-128c9404-6cdb-11e8-9126-3957ddfccb10.png)

Each aggregation tag has a `over` parameter providing you with a powerful tool that allows you to perform more complex calculations that Excel cannot do for various reasons. In particular, Excel will not be able to calculate the amount of a complex (multi-line) area. The argument to over is an expression. Example:

![tlists4_complexrange_tpl](https://user-images.githubusercontent.com/1150085/41203364-6e9ff36e-6cde-11e8-8551-671c787f7a10.png)

    <<sum over="item.AmountPaid">>
    

This function is very useful for computing subtotals in master-detail reports. You can see an example in the [Report examples](Examples#Nested-Ranges-with-Subtotals) section

Grouping
========

To perform grouping and create subtotals for the columns in ClosedXML.Report there is a tag `<<group>>`. The range is pre-sorted by all columns for which the tags are `<<group>>`, `<<sort>>`, `<<desc>>` and `<<asc>>`. The sort order for the `<<group>>` option is specified by an additional parameter - `<<desc>>` or `<<asc>>` (default asc). By default (without the use of additional options) the work of the `<<group>>` tag is similar to the work of the Subtotal method of the Range object in Excel. If the `<<group>>` tag is specified for several columns, the subtotals are grouped by all these columns. Grouping takes place from right to left, that is, first the totals are grouped by the rightmost column, for which the `<<group>>` tag is specified, then by the column with the `<<group>>` tag to the left of it, etc. The format of the service line of the range is used to format the rows of subtotals. After the subtotals are created, the service line is removed from the range.

To get subtotals you can use aggregation tags in the corresponding columns:

*   `<<Sum>>` - displays the amount by column;
*   `<<Count>>` - the number of values in the column;
*   `<<CountNums>>` - the number of non-empty values in the column;
*   `<<Avg>>` or `<<Average>>` - the average value of the column;
*   `<<Max>>` - the maximum value in the column;
*   `<<Min>>` - the minimum value in the column;
*   `<<Product>>` - product by column;
*   `<<StDev>>` - standard deviation;
*   `<<StDevP>>` - standard deviation of the total population;
*   `<<Var>>` - dispersion;
*   `<<VarP>>` - dispersion for the general population.

To change behavior, the `<<group>>` tag has a number of options:

*   Collapse - causes the subtotal to be collapsed to the level at which the `<<group>>` tag is located with this parameter.
*   MergeLabels=\[Merge1|Merge2|Merge3\] - causes the group cells to be merged in the grouped column
*   PlaceToColumn=n - allows you to specify the column in which the group header will be placed
*   WithHeader - allows you to create a group header when using subtotals
*   Disablesubtotals - allows you to disable the creation of subtotals for the column
*   DisableOutline - turns off the creation of an Outline view for the grouped column
*   PageBreaks - allows you to place each group on a separate page
*   TotalLabel - allows you to set the caption text in the line of subtotals (default: ‘Total’)
*   GrandLabel - allows you to set the caption text in the total totals line (default: ‘Grand’)

Also, to change the behavior of a grouping, there are range tags:

*   `<<SummaryAbove>>` - in case the `<<SummaryAbove>>` tag is found in the range, subtotals are placed above the data;
*   `<<DisableGrandTotal>>` - prohibits the creation of grand totals when using range grouping with subtotals.


# Tags list

| Tag | Parameters | Scope | Description |
| --- | --- | --- | --- |
| Range | source, horizontal | range | Using the `source` parameter, you can specify a data source for an range other than the name of the range. By using the `horizontal` option, you can specify that the range should be horizontally constructed. |
| Sort | asc, desc | rangeCol | Sorts the range by the column for which it is specified. The parameters `Desc` and `Asc` (Asc - by default) indicate the sort order. You can simultaneously sort by several columns. Sorting occurs from right to left, that is, first the rightmost column is sorted, then the next one is left, etc. The `<sort>` tag works both for separately located ranges, and for nested (the lowest level of nesting). |
| Asc |     | rangeCol | Same as `<sort desc>` |
| Desc |     | rangeCol | Same as `<sort asc>` |
| Group | Collapse  <br>Desc  <br>Asc  <br>MergeLabels=\[Merge1\| Merge2\| Merge3\]  <br>PlaceToColumn=n  <br>WithHeader  <br>Disablesubtotals  <br>DisableOutline  <br>PageBreaks | rangeCol | Creates subtotals on columns that have totals tags (`<sum>`, etc.), grouping them by the column for which it is specified. The range will be pre-sorted by all columns for which the Group, Sort, Desc, and Asc tags are specified. The sort order for the `<group>` tag is indicated by the optional parameter `Desc` or `Asc` (Asc by default). The `Collapse` parameter causes the intermediate totals to be rolled up to the level where the `<group>` tag is located with this parameter. The work of the tag is fully consistent with the Subtotal method of the Range object. Similarly, if the `<group>` tag is specified for multiple columns, the subtotals are grouped by all these columns. Grouping occurs from right to left, that is, first the results are grouped on the rightmost column for which the `<group>` tag is specified, then the column with the option `<group>` to the left of it, etc. To format the subtotal rows, you use the formatting of the area’s service line. After creating the subtotals, the service line is removed from the scope. The tag can be used without summary functions. In this case, the data is grouped without intermediate totals. To format the subtotal rows and group headers, formatting the area’s service line is used.  <br>The `MergeLabels` parameter causes the group cells to be combined in a grouped column.  <br>The parameter `PlaceToColumn` allows you to specify the column in which the group header will be placed.  <br>The parameter `DisableSubtotals` allows you to disable the creation of subtotals for a column.  <br>The `DisableOutline` parameter turns off the creation of the Outline view for a grouped column.  <br>The `PageBreaks` parameter allows you to put each group on a separate page.  <br>The `WithHeader` parameter allows you to create a group header when using subtotals. If the `<SummaryAbove>` tag is found (see below), the subtotals are placed above the data.  <br>For more information, see [Grouping](Grouping) |
| SummaryAbove |     | range | A helper tag for the `<group>` tag. `<SummaryAbove>` is used to place totals on groups over data. For more information, see [Grouping](Grouping) |
| DisableGrandTotal |     | range | Prevents the creation of all totals when using an area grouping with subtotals. For more information, see [Grouping](Grouping) |
| Pivot | Name  <br>Dst  <br>RowGrand  <br>ColumnGrand  <br>NoPreserveFormatting  <br>CaptionNoFormatting  <br>MergeLabels  <br>ShowButtons  <br>TreeLayout  <br>AutofitColumns  <br>NoSort |     | Creates a summary table for the range for which it is specified. The structure of the summary table is determined according to the options in the fields of the summary table (described below). The range on which the PivotTable is built must be an Excel data list. A header must be present above the scope.  <br>The `<pivot>` tag requires the mandatory specification of the `Name` parameter specifying the name of the pivot table being created. The name must be valid in Excel.  <br>The `Dst` parameter allows you to specify the exact location of the top-left corner of the pivot table (including page margins). The value of this parameter can be the cell reference formulas in styles A1 or R1C1. You can specify in this formula the name of the sheet on which you want to place the pivot table. For example, `Dst = Sheet1! D8`. If the parameter `Dst` is not specified, a new sheet with the name of this table is created for each report summary table.  <br>The parameters `RowGrand` and `ColumnGrand` include the corresponding properties of the pivot table, allowing you to get the total results on the rows and, accordingly, the columns of the table.  <br>`MergeLabels` includes the corresponding property of the PivotTable, invoking a union of cells in the row and column area.  <br>By default, all formatting of the source area is transferred to the corresponding header and data area. If the `NoPreserveFormatting` parameter is specified, the formatting is not carried over.  <br>The ShowButtons option allows you to show the expand / collapse button.  <br>`TreeLayout` - sets the pivot table mode as a tree.  <br>`AutofitColumns` - includes auto-match the width of the columns of the pivot table.  <br>`NoSort` - disables automatic sorting of the pivot table.  <br>  <br>Pivot table tags (Page, Row, Column, Data) are used to describe the structure of the pivot table, defining the areas of the pivot table into which these fields are placed. In conjunction with them, the final tags are used, describing the types of totals that must be obtained from these fields. Columns for which the PivotTable field tags are not specified are included in the summary table as hidden. In the process of working with a finished report using standard Excel tools, it is possible to modify the structure of the summary table by the end user. The structure of the PivotTable should take into account the restrictions on the summary tables, which are described in the documentation for the specific version of Excel. A detailed description of the methodology for calculating these constraints can be found in MSDN.  <br>  <br>More information in the section [Pivot tables](Pivot-tables) |
| Page |     | rangeCol | This tag places the column for which it is specified in the field of the pages in the pivot table. The field label is a cell value from the table header above the source area |
| Row |     | rangeCol | This tag places the column for which it is specified in the area of ​​the rows in the pivot table. The field label is the cell value from the table header above the source area. Excel automatically groups the elements of the internal field for each element of the external field of the rows and, if necessary, creates subtotals. The type of subtotals is determined by an additional specify of the totals tag for the column with the tag `<row>`. For example, specifying the tags `<row> <sum> <count>` will create an amount and a count in the intermediate totals by the field of the pivot table, the source of which is the column with these tags. |
| Column, Col |     | rangeCol | This tag places the column for which it is specified in the column area of the pivot table. The field label is the cell value from the table header above the source area. You can get one or more subtotals on column fields. Their type is determined by an additional indication of the totals tag for the column with the tag `<col>`. For example, specifying the options `<col> <sum> <count>` will create an amount and a count in the intermediate totals by the field of the pivot table, the source of which is the column with these tags. |
| Data |     | rangeCol | This tag places the column for which it is specified in the pivot table data area. The field label is the cell value from the table header above the source area. By default, the `sum` function is applied to the data fields. To switch to any other summary function in the data field, jointly with the tag `<data>` you can use the totals tags. |
| SUM  <br>AVG  <br>AVERAGE  <br>COUNT  <br>COUNTNUMS  <br>MAX  <br>MIN  <br>PRODUCT  <br>STDEV  <br>STDEVP  <br>VAR  <br>VARP | over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression”  <br>over=”expression” | rangeCol | These are the tags of the totals. Specifying it for a column causes a summary of the column. The result is calculated by the corresponding Excel function. The results are placed in the service column of the range. On the column, you can get only one kind of total (for example, the amount or only the average). If several summary options are specified, only the last specified total is calculated. The `over` parameter allows you to make calculations using the power of .NET and LINQ. Expressions use the syntax of C# expressions. Using the `items` variable, you can access a list of table items. |
| OnlyValues |     | worksheet  <br>range  <br>rangeCol  <br>cell | Replaces all formulas on the worksheet, in the region, in the column of the region, or in the cell where it is specified, by the values of these formulas. |
| AutoFilter |     | range | Enables the AutoFilter in the area for which it is specified. |
| Protected | Password=”password” | workbook  <br>worksheetrange  <br>cell  <br>rangeCol | If specified for the report, it protects all sheets and protects the book itself. If specified for a sheet, this sheet is protected. Analagically applied to the area, a certain column of the region and to a separate cell. The `Password` option is optional. If the password is not set, it will be generated randomly. |
| ColsFit |     | workbook  <br>worksheet  <br>worksheetCol  <br>range  <br>rangeCol  <br>cell | Causes automatic alignment of column widths by value in the cells of the columns of the entire report (if specified for the report), the sheet (if specified for the sheet), the entire sheet column (if specified on the first line of the sheet), the range (if specified for the range), the range column (if specified in the options line) or for one cell. If specified for the sheet and for the range on this sheet, it is only called for the sheet. The same principle is applied from the greater to the less for the remaining cases. |
| RowsFit |     | workbook  <br>worksheet  <br>worksheetRow  <br>range  <br>rangeCol  <br>cell | Causes automatic alignment of row height by value in the cells of the rows of the entire report (if specified for the report), the sheet (if specified for the sheet), the entire row of the sheet (if specified on the first column of the sheet), the range (if specified for the range) or for one cell. If specified for the sheet and for the range on this sheet, it is only called for the sheet. The same principle is applied from the greater to the less for the remaining cases. |
| Hidden, Hide |     | worksheet | Hides the sheet for which it is specified. In the case of debugging the report, produces a “soft” hiding, which makes the sheet visible. |
| Delete | disabled=`<string>` | worksheet  <br>worksheetRow  <br>worksheetCol  <br>rangeCol | Removes a sheet, a row/column of a sheet, or an range column, depending on where the tag is located.  <br>The `disabled` option allows to disable `Delete` tag execution. Deletion will be performed if the tag is empty or has one of the values: `false`, `ложь`, `no`, `not`, `null`, `0`, `0.0`, `0,0`, `-`. In other cases of applying the `disabled` parameter, the tag will be ignored. |
| PageOptions | Wide=`<int>`  <br>Tall=`<int>`  <br>Landscape | workbook  <br>worksheet | Specifies the page parameters for printing, adjusting the width of the sheet with the parameter `Wide`, the height of the sheet with the parameter `Tall` and the orientation of the sheet with the parameter `Landscape` |
| Height | size, example `50` |     | Sets the row height. |
| HeightRange | size, example `50` |     | Sets the row height for the data source in the area. All rows within the area will have their height adjusted. |

