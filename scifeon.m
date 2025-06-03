let
    GetNamedValue = (valueName as text, defaultValue as text) as text =>
        let
            namedItem = try Excel.CurrentWorkbook(){[Name=valueName]}[Content] otherwise null,
            rawValue = 
                if namedItem <> null and Table.RowCount(namedItem) > 0 then 
                    try namedItem{0}[Column1] otherwise defaultValue
                else 
                    defaultValue,
            value = if rawValue <> null then Text.From(rawValue) else defaultValue
        in
            value,
    EscapeODataValue = (value as text) => Text.Replace(value, "'", "''"),
    
    Eq = (col as text, val as text) => col & " eq '" & EscapeODataValue(val) & "'",
    Ne = (col as text, val as text) => col & " ne '" & EscapeODataValue(val) & "'",
    Ge = (col as text, val as text) => col & " ge '" & EscapeODataValue(val) & "'",
    Gt = (col as text, val as text) => col & " gt '" & EscapeODataValue(val) & "'",
    Le = (col as text, val as text) => col & " le '" & EscapeODataValue(val) & "'",
    Lt = (col as text, val as text) => col & " ge '" & EscapeODataValue(val) & "'",
    In = (col as text, vals as list) => col & " in (" & Text.Combine(List.Transform(vals, each "'" & EscapeODataValue(_) & "'"), ", ") & ")",
    NotIn = (col as text, vals as list) => col & " not_in (" & Text.Combine(List.Transform(vals, each "'" & EscapeODataValue(_) & "'"), ", ") & ")",
    IsNull = (col as text) => col & " eq null",
    IsNotNull = (col as text) => col & " ne null",
    IsEmpty = (col as text) => "(" & col & " eq null or " & col & " eq '')",
    IsNotEmpty = (col as text) => "(" & col & " ne null and " & col & " ne '')",
    Contains = (col as text, val as text) => "contains(" & col & ", '" & EscapeODataValue(val) & "')",
    DoesNotContain = (col as text, val as text) => "not contains(" & col & ", '" & EscapeODataValue(val) & "')",
    StartsWith = (col as text, val as text) => "startswith(" & col & ", '" & EscapeODataValue(val) & "')",
    EndsWith = (col as text, val as text) => "endswith(" & col & ", '" & EscapeODataValue(val) & "')",
    Template = (col as text, val as text) => col & " template '" & EscapeODataValue(val) & "'",
    And = (conds as list) => Text.Combine(conds, " and "),
    Or = (conds as list) => "(" & Text.Combine(conds, " or ") & ")",
    Query = (
        InstanceUrl as text,
        View as text,
        optional Select as list,
        optional Filters as list,
        optional OrderBy as nullable text
    ) as table =>
        let
            // Normalize InstanceUrl
            BaseUrl = if Text.EndsWith(InstanceUrl, "/") then InstanceUrl else InstanceUrl & "/",
            SelectPart = if Select <> null and List.Count(Select) > 0 then "$select=" & Text.Combine(Select, ",") else null,
            FilterPart = if Filters <> null and List.Count(Filters) > 0 then "$filter=" & And(Filters) else null,
            OrderByPart = if OrderBy <> null and Text.Length(OrderBy) > 0 then "$orderby=" & OrderBy else null,

            QueryParams = List.Select({FilterPart, OrderByPart, SelectPart}, each _ <> null),
            QueryString = Text.Combine(QueryParams, "&"),
            FullUrl = BaseUrl & "odata/" & View & (if Text.Length(QueryString) > 0 then "?" & QueryString else ""),

            Source = OData.Feed(FullUrl, null, [Implementation="2.0"]),
            Output = if Select <> null and List.Count(Select) > 0 then Table.SelectColumns(Source, Select) else Source
        in
            Output,
    Filter = [
        Eq = Eq, Ne = Ne, Ge = Ge, Gt = Gt,
        Le = Le, Lt = Lt,
        In = In, NotIn = NotIn,
        IsNull = IsNull, IsNotNull = IsNotNull,
        IsEmpty = IsEmpty, IsNotEmpty = IsNotEmpty,
        Contains = Contains, DoesNotContain = DoesNotContain,
        StartsWith = StartsWith, EndsWith = EndsWith,
        Template = Template,
        And = And, Or = Or
    ]
in
    [
        GetNamedValue = GetNamedValue,
        Filter = Filter,
        Query = Query
    ]
