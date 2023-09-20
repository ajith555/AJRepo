let
    Source = ...  // Connect to your Exchange Online data source
    Inbox = Source{[FolderName="Inbox"]}[Data],
    SentItems = Source{[FolderName="Sent Items"]}[Data],

    // Filter Sent Items for responses
    FilteredSentItems = Table.SelectColumns(SentItems, {"Id", "DateTimeSent"}),
    
    // Merge Inbox with filtered Sent Items using the "Id" column
    Merged = Table.NestedJoin(Inbox, "Id", FilteredSentItems, "Id", "ResponseData", JoinKind.LeftOuter),
    
    // Expand the merged table to get response date
    Expanded = Table.ExpandTableColumn(Merged, "ResponseData", {"DateTimeSent"}, {"ResponseDateTimeSent"}),
    
    // Calculate the response time
    CustomColumn = Table.AddColumn(Expanded, "ResponseTime", each [ResponseDateTimeSent] - [DateTimeSent]),
    
    // Extract the week number
    WeekNumber = Table.AddColumn(CustomColumn, "WeekNumber", each Date.WeekOfYear([DateTimeSent])),
    
    // Group by week and calculate average response time
    Grouped = Table.Group(WeekNumber, "WeekNumber", {{"AvgResponseTime", each List.Average([ResponseTime])}})
in
    Grouped
