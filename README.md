# Scifeon OData Utils for Power Query

1. Open the Power Query Editor
2. Create a new Blank Query
3. Name the query "Scifeon"
4. Open the Advanced Editor
5. Copy and paste the contents of `scifeon.m` in the editor

<img width="903" alt="image" src="https://github.com/user-attachments/assets/f92430c2-7d25-400c-bd67-ff714f02ed4a" />


## Example Query
Example query that fetches all samples that were created within the current year.

```
let
    GetNamedValue = Scifeon[GetNamedValue],
    Filter = Scifeon[Filter],
    Query = Scifeon[Query],

    InstanceUrl = GetNamedValue("instance_url", "https://my-scifeon-url.scifeon.cloud"),
    CurrentYear = Date.Year(DateTime.LocalNow()),
    YearFrom = GetNamedValue("year_from", Text.From(CurrentYear)),
    YearTo = GetNamedValue("year_to", Text.From(CurrentYear)),
    View = "Sample",
    Select = { "ID", "Name", "Description", "Type", "CreatedBy", "CreatedUtc"},
    Filters = { 
        Filter[NotIn]("status", { "Discarded", "Deleted", "Canceled"}),
        Filter[Ge]("createdUtc", YearFrom & "-01-01T00:00:00"),
        Filter[Le]("CreatedUtc", YearTo & "-12-31T00:00:00")
        },
    OrderBy = "dato asc",

    QueryResult = Query(InstanceUrl, View, Select, Filters, OrderBy),
in
    QueryResult
```

## Table of Contents

* [GetNamedValue](#getnamedvalue)
* [Filter Functions](#filter-functions)
  * [Comparison](#comparison)
  * [Null Checks](#null-checks)
  * [String Operations](#string-operations)
  * [Logical Combinators](#logical-combinators)
* [Query](#query)

## GetNamedValue

**Description**
Retrieves a named value (from a named cell) in the current Excel workbook. Returns a default value if the name is not found or is empty.

**Signature**
GetNamedValue(valueName as text, defaultValue as text) as text

**Parameters**

| Name         | Type | Description                                    |
| ------------ | ---- | ---------------------------------------------- |
| valueName    | text | The name of the named item in the workbook.    |
| defaultValue | text | A fallback value if the named item is missing. |

**Returns**
A text value from the named expression, or the default value.

**Example**
`GetNamedValue("instance_url", "https://some-backup-url.scifeon.cloud")`


## Filter Functions

These are used to construct OData `$filter` expressions.

### Comparison

* Equal: `Eq(col as text, val as text)`
* Not Equal: `Ne(col as text, val as text)`
* Greater or Equal: `Ge(col as text, val as text)`
* Greater Than: `Gt(col as text, val as text)`
* Less or Equal: `Le(col as text, val as text)`
* Less Than: `Lt(col as text, val as text)`

**Example**
`Filter[Eq]("Status", "Open")` → "Status eq 'Open'"


### List Operators

* `In(col as text, vals as list)`
* `NotIn(col as text, vals as list)`

**Example**
`Filter[In]("Category", {"A", "B"})` → "Category in ('A', 'B')"

### Null Checks

* `IsNull(col as text)`
* `IsNotNull(col as text)`
* `IsEmpty(col as text)`
* `IsNotEmpty(col as text)`

**Example**
`Filter[IsEmpty]("Comment")` → "Comment isempty"


### String Operations

* Contains(col as text, val as text)
* DoesNotContain(col as text, val as text)
* StartsWith(col as text, val as text)
* EndsWith(col as text, val as text)
* Template(col as text, val as text)

**Example**
`Filter[Contains]("Title", "Report")` → "contains(Title, 'Report')"

### Logical Combinators

* `And(conds as list)`
* `Or(conds as list)`

**Example**
`Filter[And]({
  Filter[Eq]("Status", "Open"),
  Filter[IsNotNull]("Assignee")
})`


## Query

**Description**
Builds and executes an OData query against a given instance and view, supporting filters, selection, and ordering.

**Signature**
Query(
 InstanceUrl as text,
 View as text,
 optional Select as list,
 optional Filters as list,
 optional OrderBy as nullable text
) as table

**Parameters**

| Name        | Type          | Description                                                   |
| ----------- | ------------- | ------------------------------------------------------------- |
| InstanceUrl | text          | Base URL of the OData service.                                |
| View        | text          | Name of the OData entity/view to query.                       |
| Select      | optional list | Columns to return. Use Uppercase Column Names from Datamodel  |
| Filters     | optional list | List of filter expressions, use Filter helpers.               |
| OrderBy     | optional text | Column to order by (OData syntax).                            |

**Returns**
A table resulting from the OData query.

**Example**
```
Query(
 "https://some-scifeon-url.scifeon.cloud",
 "Sample",
 {"Id", "Name", "Status"},
 {
  Filter[Eq]("Status", "Open"),
  Filter[StartsWith]("Name", "Test_")
 },
 "Title asc"
)
```

Let me know if you want this saved as a `.md` file or integrated with comments directly inside your code.
