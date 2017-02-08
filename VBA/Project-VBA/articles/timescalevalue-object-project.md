---
title: TimeScaleValue Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.TimeScaleValue
ms.assetid: bea0ad82-a3de-30d8-f191-dc2248c32653
---


# TimeScaleValue Object (Project)

Represents a timescaled data item. The  **TimeScaleValue** object is a member of the **[TimeScaleValues](timescalevalues-object-project.md)** collection.


## Examples

 **Using the TimeScaleValue Object**

Use  **TimeScaleValues** ( _Index_ ), where _Index_ is the index number of the timescaled data item, to return a single **TimeScaleValue** object. The following example displays the number of hours of work per day for a resource during the first full week in October 2012.




```
Dim TSV As TimeScaleValues, HowMany As Long
Dim HoursPerDay As String

Set TSV = ActiveCell.Resource.TimeScaleData("10/1/2012", "10/5/2012", TimescaleUnit:=pjTimescaleDays)

For HowMany = 1 To TSV.Count
    HoursPerDay = HoursPerDay &amp; TSV(HowMany).StartDate &amp; " - " &amp; _
        TSV(HowMany).EndDate &amp; ", " &amp; TSV(HowMany) / 60 &amp; vbCrLf
Next HowMany

MsgBox HoursPerDay
```

 **Using the TimeScaleValues Collection**

Use the  **[TimeScaleData](http://msdn.microsoft.com/library/resource-timescaledata-method-project%28Office.15%29.aspx)** method to return a **TimeScaleValues** collection. The following example returns a **TimeScaleValues** collection for the amount of work done by the resource in the active cell between the specified dates, split into week-long portions.




```
ActiveCell.Resource.TimeScaleData("10/1/2012", "10/31/2012")
```

Use the  **[Add](http://msdn.microsoft.com/library/timescalevalues-add-method-project%28Office.15%29.aspx)** method to add a **TimeScaleValue** object to the **TimeScaleValues** collection. The following example adds 8 hours of work to Tuesday of that week.




```
Dim TSV As TimeScaleValues

Set TSV = ActiveCell.Resource.TimeScaleData("10/1/2012", "10/5/2012", TimescaleUnit:=pjTimescaleDays)
TSV.Add 480, 2
```


## Methods



|**Name**|
|:-----|
|[Clear](http://msdn.microsoft.com/library/timescalevalue-clear-method-project%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/timescalevalue-delete-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/timescalevalue-application-property-project%28Office.15%29.aspx)|
|[EndDate](http://msdn.microsoft.com/library/timescalevalue-enddate-property-project%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/timescalevalue-index-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/timescalevalue-parent-property-project%28Office.15%29.aspx)|
|[StartDate](http://msdn.microsoft.com/library/timescalevalue-startdate-property-project%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/timescalevalue-value-property-project%28Office.15%29.aspx)|

