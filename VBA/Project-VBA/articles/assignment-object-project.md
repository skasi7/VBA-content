---
title: Assignment Object (Project)
ms.prod: PROJECTSERVER
api_name:
- Project.Assignment
ms.assetid: bfb9a505-7818-0a86-9d4b-f19a0ff465d3
---


# Assignment Object (Project)

Represents an assignment for a task or resource. The  **Assignment** object is a member of an **[Assignments](assignments-object-project.md)** or an **[OverAllocatedAssignments](http://msdn.microsoft.com/library/overallocatedassignments-object-project%28Office.15%29.aspx)** collection.


## Example

 **Using the Assignment Object**

Use  **Assignments** ( _Index_ ), where _Index_ is the assignment index number, to return a single **Assignment** object. The following example displays the name of the first resource assigned to the specified task.




```
MsgBox ActiveProject.Tasks(1).Assignments(1).ResourceName
```

 **Using the Assignments Collection**

Use the  **[Assignments](http://msdn.microsoft.com/library/task-assignments-property-project%28Office.15%29.aspx)** property to return an **Assignments** collection. The following example displays all the resources assigned to the specified task.




```
Dim A As Assignment 
 
For Each A In ActiveProject.Tasks(1).Assignments 
 MsgBox A.ResourceName 
Next A
```

Use the  **[Add](http://msdn.microsoft.com/library/assignments-add-method-project%28Office.15%29.aspx)** method to add an **Assignment** object to the **Assignments** collection. The following example adds a resource identified by the number 212 as a new assignment for the specified task.




```
ActiveProject.Tasks(1).Assignments.Add ResourceID:=212
```


## Methods



|**Name**|
|:-----|
|[AppendNotes](http://msdn.microsoft.com/library/assignment-appendnotes-method-project%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/assignment-delete-method-project%28Office.15%29.aspx)|
|[EnterpriseTeamMember](http://msdn.microsoft.com/library/assignment-enterpriseteammember-method-project%28Office.15%29.aspx)|
|[Replan](http://msdn.microsoft.com/library/assignment-replan-method-project%28Office.15%29.aspx)|
|[TimeScaleData](http://msdn.microsoft.com/library/assignment-timescaledata-method-project%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActualCost](http://msdn.microsoft.com/library/assignment-actualcost-property-project%28Office.15%29.aspx)|
|[ActualFinish](http://msdn.microsoft.com/library/assignment-actualfinish-property-project%28Office.15%29.aspx)|
|[ActualOvertimeCost](http://msdn.microsoft.com/library/assignment-actualovertimecost-property-project%28Office.15%29.aspx)|
|[ActualOvertimeWork](http://msdn.microsoft.com/library/assignment-actualovertimework-property-project%28Office.15%29.aspx)|
|[ActualStart](http://msdn.microsoft.com/library/assignment-actualstart-property-project%28Office.15%29.aspx)|
|[ActualWork](http://msdn.microsoft.com/library/assignment-actualwork-property-project%28Office.15%29.aspx)|
|[ACWP](http://msdn.microsoft.com/library/assignment-acwp-property-project%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/assignment-application-property-project%28Office.15%29.aspx)|
|[Baseline10BudgetCost](http://msdn.microsoft.com/library/assignment-baseline10budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline10BudgetWork](http://msdn.microsoft.com/library/assignment-baseline10budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline10Cost](http://msdn.microsoft.com/library/assignment-baseline10cost-property-project%28Office.15%29.aspx)|
|[Baseline10Finish](http://msdn.microsoft.com/library/assignment-baseline10finish-property-project%28Office.15%29.aspx)|
|[Baseline10Start](http://msdn.microsoft.com/library/assignment-baseline10start-property-project%28Office.15%29.aspx)|
|[Baseline10Work](http://msdn.microsoft.com/library/assignment-baseline10work-property-project%28Office.15%29.aspx)|
|[Baseline1BudgetCost](http://msdn.microsoft.com/library/assignment-baseline1budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline1BudgetWork](http://msdn.microsoft.com/library/assignment-baseline1budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline1Cost](http://msdn.microsoft.com/library/assignment-baseline1cost-property-project%28Office.15%29.aspx)|
|[Baseline1Finish](http://msdn.microsoft.com/library/assignment-baseline1finish-property-project%28Office.15%29.aspx)|
|[Baseline1Start](http://msdn.microsoft.com/library/assignment-baseline1start-property-project%28Office.15%29.aspx)|
|[Baseline1Work](http://msdn.microsoft.com/library/assignment-baseline1work-property-project%28Office.15%29.aspx)|
|[Baseline2BudgetCost](http://msdn.microsoft.com/library/assignment-baseline2budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline2BudgetWork](http://msdn.microsoft.com/library/assignment-baseline2budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline2Cost](http://msdn.microsoft.com/library/assignment-baseline2cost-property-project%28Office.15%29.aspx)|
|[Baseline2Finish](http://msdn.microsoft.com/library/assignment-baseline2finish-property-project%28Office.15%29.aspx)|
|[Baseline2Start](http://msdn.microsoft.com/library/assignment-baseline2start-property-project%28Office.15%29.aspx)|
|[Baseline2Work](http://msdn.microsoft.com/library/assignment-baseline2work-property-project%28Office.15%29.aspx)|
|[Baseline3BudgetCost](http://msdn.microsoft.com/library/assignment-baseline3budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline3BudgetWork](http://msdn.microsoft.com/library/assignment-baseline3budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline3Cost](http://msdn.microsoft.com/library/assignment-baseline3cost-property-project%28Office.15%29.aspx)|
|[Baseline3Finish](http://msdn.microsoft.com/library/assignment-baseline3finish-property-project%28Office.15%29.aspx)|
|[Baseline3Start](http://msdn.microsoft.com/library/assignment-baseline3start-property-project%28Office.15%29.aspx)|
|[Baseline3Work](http://msdn.microsoft.com/library/assignment-baseline3work-property-project%28Office.15%29.aspx)|
|[Baseline4BudgetCost](http://msdn.microsoft.com/library/assignment-baseline4budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline4BudgetWork](http://msdn.microsoft.com/library/assignment-baseline4budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline4Cost](http://msdn.microsoft.com/library/assignment-baseline4cost-property-project%28Office.15%29.aspx)|
|[Baseline4Finish](http://msdn.microsoft.com/library/assignment-baseline4finish-property-project%28Office.15%29.aspx)|
|[Baseline4Start](http://msdn.microsoft.com/library/assignment-baseline4start-property-project%28Office.15%29.aspx)|
|[Baseline4Work](http://msdn.microsoft.com/library/assignment-baseline4work-property-project%28Office.15%29.aspx)|
|[Baseline5BudgetCost](http://msdn.microsoft.com/library/assignment-baseline5budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline5BudgetWork](http://msdn.microsoft.com/library/assignment-baseline5budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline5Cost](http://msdn.microsoft.com/library/assignment-baseline5cost-property-project%28Office.15%29.aspx)|
|[Baseline5Finish](http://msdn.microsoft.com/library/assignment-baseline5finish-property-project%28Office.15%29.aspx)|
|[Baseline5Start](http://msdn.microsoft.com/library/assignment-baseline5start-property-project%28Office.15%29.aspx)|
|[Baseline5Work](http://msdn.microsoft.com/library/assignment-baseline5work-property-project%28Office.15%29.aspx)|
|[Baseline6BudgetCost](http://msdn.microsoft.com/library/assignment-baseline6budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline6BudgetWork](http://msdn.microsoft.com/library/assignment-baseline6budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline6Cost](http://msdn.microsoft.com/library/assignment-baseline6cost-property-project%28Office.15%29.aspx)|
|[Baseline6Finish](http://msdn.microsoft.com/library/assignment-baseline6finish-property-project%28Office.15%29.aspx)|
|[Baseline6Start](http://msdn.microsoft.com/library/assignment-baseline6start-property-project%28Office.15%29.aspx)|
|[Baseline6Work](http://msdn.microsoft.com/library/assignment-baseline6work-property-project%28Office.15%29.aspx)|
|[Baseline7BudgetCost](http://msdn.microsoft.com/library/assignment-baseline7budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline7BudgetWork](http://msdn.microsoft.com/library/assignment-baseline7budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline7Cost](http://msdn.microsoft.com/library/assignment-baseline7cost-property-project%28Office.15%29.aspx)|
|[Baseline7Finish](http://msdn.microsoft.com/library/assignment-baseline7finish-property-project%28Office.15%29.aspx)|
|[Baseline7Start](http://msdn.microsoft.com/library/assignment-baseline7start-property-project%28Office.15%29.aspx)|
|[Baseline7Work](http://msdn.microsoft.com/library/assignment-baseline7work-property-project%28Office.15%29.aspx)|
|[Baseline8BudgetCost](http://msdn.microsoft.com/library/assignment-baseline8budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline8BudgetWork](http://msdn.microsoft.com/library/assignment-baseline8budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline8Cost](http://msdn.microsoft.com/library/assignment-baseline8cost-property-project%28Office.15%29.aspx)|
|[Baseline8Finish](http://msdn.microsoft.com/library/assignment-baseline8finish-property-project%28Office.15%29.aspx)|
|[Baseline8Start](http://msdn.microsoft.com/library/assignment-baseline8start-property-project%28Office.15%29.aspx)|
|[Baseline8Work](http://msdn.microsoft.com/library/assignment-baseline8work-property-project%28Office.15%29.aspx)|
|[Baseline9BudgetCost](http://msdn.microsoft.com/library/assignment-baseline9budgetcost-property-project%28Office.15%29.aspx)|
|[Baseline9BudgetWork](http://msdn.microsoft.com/library/assignment-baseline9budgetwork-property-project%28Office.15%29.aspx)|
|[Baseline9Cost](http://msdn.microsoft.com/library/assignment-baseline9cost-property-project%28Office.15%29.aspx)|
|[Baseline9Finish](http://msdn.microsoft.com/library/assignment-baseline9finish-property-project%28Office.15%29.aspx)|
|[Baseline9Start](http://msdn.microsoft.com/library/assignment-baseline9start-property-project%28Office.15%29.aspx)|
|[Baseline9Work](http://msdn.microsoft.com/library/assignment-baseline9work-property-project%28Office.15%29.aspx)|
|[BaselineBudgetCost](http://msdn.microsoft.com/library/assignment-baselinebudgetcost-property-project%28Office.15%29.aspx)|
|[BaselineBudgetWork](http://msdn.microsoft.com/library/assignment-baselinebudgetwork-property-project%28Office.15%29.aspx)|
|[BaselineCost](http://msdn.microsoft.com/library/assignment-baselinecost-property-project%28Office.15%29.aspx)|
|[BaselineFinish](http://msdn.microsoft.com/library/assignment-baselinefinish-property-project%28Office.15%29.aspx)|
|[BaselineStart](http://msdn.microsoft.com/library/assignment-baselinestart-property-project%28Office.15%29.aspx)|
|[BaselineWork](http://msdn.microsoft.com/library/assignment-baselinework-property-project%28Office.15%29.aspx)|
|[BCWP](http://msdn.microsoft.com/library/assignment-bcwp-property-project%28Office.15%29.aspx)|
|[BCWS](http://msdn.microsoft.com/library/assignment-bcws-property-project%28Office.15%29.aspx)|
|[BookingType](http://msdn.microsoft.com/library/assignment-bookingtype-property-project%28Office.15%29.aspx)|
|[BudgetCost](http://msdn.microsoft.com/library/assignment-budgetcost-property-project%28Office.15%29.aspx)|
|[BudgetWork](http://msdn.microsoft.com/library/assignment-budgetwork-property-project%28Office.15%29.aspx)|
|[Confirmed](http://msdn.microsoft.com/library/assignment-confirmed-property-project%28Office.15%29.aspx)|
|[Cost](http://msdn.microsoft.com/library/assignment-cost-property-project%28Office.15%29.aspx)|
|[Cost1](http://msdn.microsoft.com/library/assignment-cost1-property-project%28Office.15%29.aspx)|
|[Cost10](http://msdn.microsoft.com/library/assignment-cost10-property-project%28Office.15%29.aspx)|
|[Cost2](http://msdn.microsoft.com/library/assignment-cost2-property-project%28Office.15%29.aspx)|
|[Cost3](http://msdn.microsoft.com/library/assignment-cost3-property-project%28Office.15%29.aspx)|
|[Cost4](http://msdn.microsoft.com/library/assignment-cost4-property-project%28Office.15%29.aspx)|
|[Cost5](http://msdn.microsoft.com/library/assignment-cost5-property-project%28Office.15%29.aspx)|
|[Cost6](http://msdn.microsoft.com/library/assignment-cost6-property-project%28Office.15%29.aspx)|
|[Cost7](http://msdn.microsoft.com/library/assignment-cost7-property-project%28Office.15%29.aspx)|
|[Cost8](http://msdn.microsoft.com/library/assignment-cost8-property-project%28Office.15%29.aspx)|
|[Cost9](http://msdn.microsoft.com/library/assignment-cost9-property-project%28Office.15%29.aspx)|
|[CostRateTable](http://msdn.microsoft.com/library/assignment-costratetable-property-project%28Office.15%29.aspx)|
|[CostVariance](http://msdn.microsoft.com/library/assignment-costvariance-property-project%28Office.15%29.aspx)|
|[Created](http://msdn.microsoft.com/library/assignment-created-property-project%28Office.15%29.aspx)|
|[CV](http://msdn.microsoft.com/library/assignment-cv-property-project%28Office.15%29.aspx)|
|[Date1](http://msdn.microsoft.com/library/assignment-date1-property-project%28Office.15%29.aspx)|
|[Date10](http://msdn.microsoft.com/library/assignment-date10-property-project%28Office.15%29.aspx)|
|[Date2](http://msdn.microsoft.com/library/assignment-date2-property-project%28Office.15%29.aspx)|
|[Date3](http://msdn.microsoft.com/library/assignment-date3-property-project%28Office.15%29.aspx)|
|[Date4](http://msdn.microsoft.com/library/assignment-date4-property-project%28Office.15%29.aspx)|
|[Date5](http://msdn.microsoft.com/library/assignment-date5-property-project%28Office.15%29.aspx)|
|[Date6](http://msdn.microsoft.com/library/assignment-date6-property-project%28Office.15%29.aspx)|
|[Date7](http://msdn.microsoft.com/library/assignment-date7-property-project%28Office.15%29.aspx)|
|[Date8](http://msdn.microsoft.com/library/assignment-date8-property-project%28Office.15%29.aspx)|
|[Date9](http://msdn.microsoft.com/library/assignment-date9-property-project%28Office.15%29.aspx)|
|[Delay](http://msdn.microsoft.com/library/assignment-delay-property-project%28Office.15%29.aspx)|
|[Duration1](http://msdn.microsoft.com/library/assignment-duration1-property-project%28Office.15%29.aspx)|
|[Duration10](http://msdn.microsoft.com/library/assignment-duration10-property-project%28Office.15%29.aspx)|
|[Duration2](http://msdn.microsoft.com/library/assignment-duration2-property-project%28Office.15%29.aspx)|
|[Duration3](http://msdn.microsoft.com/library/assignment-duration3-property-project%28Office.15%29.aspx)|
|[Duration4](http://msdn.microsoft.com/library/assignment-duration4-property-project%28Office.15%29.aspx)|
|[Duration5](http://msdn.microsoft.com/library/assignment-duration5-property-project%28Office.15%29.aspx)|
|[Duration6](http://msdn.microsoft.com/library/assignment-duration6-property-project%28Office.15%29.aspx)|
|[Duration7](http://msdn.microsoft.com/library/assignment-duration7-property-project%28Office.15%29.aspx)|
|[Duration8](http://msdn.microsoft.com/library/assignment-duration8-property-project%28Office.15%29.aspx)|
|[Duration9](http://msdn.microsoft.com/library/assignment-duration9-property-project%28Office.15%29.aspx)|
|[Finish](http://msdn.microsoft.com/library/assignment-finish-property-project%28Office.15%29.aspx)|
|[Finish1](http://msdn.microsoft.com/library/assignment-finish1-property-project%28Office.15%29.aspx)|
|[Finish10](http://msdn.microsoft.com/library/assignment-finish10-property-project%28Office.15%29.aspx)|
|[Finish2](http://msdn.microsoft.com/library/assignment-finish2-property-project%28Office.15%29.aspx)|
|[Finish3](http://msdn.microsoft.com/library/assignment-finish3-property-project%28Office.15%29.aspx)|
|[Finish4](http://msdn.microsoft.com/library/assignment-finish4-property-project%28Office.15%29.aspx)|
|[Finish5](http://msdn.microsoft.com/library/assignment-finish5-property-project%28Office.15%29.aspx)|
|[Finish6](http://msdn.microsoft.com/library/assignment-finish6-property-project%28Office.15%29.aspx)|
|[Finish7](http://msdn.microsoft.com/library/assignment-finish7-property-project%28Office.15%29.aspx)|
|[Finish8](http://msdn.microsoft.com/library/assignment-finish8-property-project%28Office.15%29.aspx)|
|[Finish9](http://msdn.microsoft.com/library/assignment-finish9-property-project%28Office.15%29.aspx)|
|[FinishVariance](http://msdn.microsoft.com/library/assignment-finishvariance-property-project%28Office.15%29.aspx)|
|[FixedMaterialAssignment](http://msdn.microsoft.com/library/assignment-fixedmaterialassignment-property-project%28Office.15%29.aspx)|
|[Flag1](http://msdn.microsoft.com/library/assignment-flag1-property-project%28Office.15%29.aspx)|
|[Flag10](http://msdn.microsoft.com/library/assignment-flag10-property-project%28Office.15%29.aspx)|
|[Flag11](http://msdn.microsoft.com/library/assignment-flag11-property-project%28Office.15%29.aspx)|
|[Flag12](http://msdn.microsoft.com/library/assignment-flag12-property-project%28Office.15%29.aspx)|
|[Flag13](http://msdn.microsoft.com/library/assignment-flag13-property-project%28Office.15%29.aspx)|
|[Flag14](http://msdn.microsoft.com/library/assignment-flag14-property-project%28Office.15%29.aspx)|
|[Flag15](http://msdn.microsoft.com/library/assignment-flag15-property-project%28Office.15%29.aspx)|
|[Flag16](http://msdn.microsoft.com/library/assignment-flag16-property-project%28Office.15%29.aspx)|
|[Flag17](http://msdn.microsoft.com/library/assignment-flag17-property-project%28Office.15%29.aspx)|
|[Flag18](http://msdn.microsoft.com/library/assignment-flag18-property-project%28Office.15%29.aspx)|
|[Flag19](http://msdn.microsoft.com/library/assignment-flag19-property-project%28Office.15%29.aspx)|
|[Flag2](http://msdn.microsoft.com/library/assignment-flag2-property-project%28Office.15%29.aspx)|
|[Flag20](http://msdn.microsoft.com/library/assignment-flag20-property-project%28Office.15%29.aspx)|
|[Flag3](http://msdn.microsoft.com/library/assignment-flag3-property-project%28Office.15%29.aspx)|
|[Flag4](http://msdn.microsoft.com/library/assignment-flag4-property-project%28Office.15%29.aspx)|
|[Flag5](http://msdn.microsoft.com/library/assignment-flag5-property-project%28Office.15%29.aspx)|
|[Flag6](http://msdn.microsoft.com/library/assignment-flag6-property-project%28Office.15%29.aspx)|
|[Flag7](http://msdn.microsoft.com/library/assignment-flag7-property-project%28Office.15%29.aspx)|
|[Flag8](http://msdn.microsoft.com/library/assignment-flag8-property-project%28Office.15%29.aspx)|
|[Flag9](http://msdn.microsoft.com/library/assignment-flag9-property-project%28Office.15%29.aspx)|
|[Guid](http://msdn.microsoft.com/library/assignment-guid-property-project%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/assignment-hyperlink-property-project%28Office.15%29.aspx)|
|[HyperlinkAddress](http://msdn.microsoft.com/library/assignment-hyperlinkaddress-property-project%28Office.15%29.aspx)|
|[HyperlinkHREF](http://msdn.microsoft.com/library/assignment-hyperlinkhref-property-project%28Office.15%29.aspx)|
|[HyperlinkScreenTip](http://msdn.microsoft.com/library/assignment-hyperlinkscreentip-property-project%28Office.15%29.aspx)|
|[HyperlinkSubAddress](http://msdn.microsoft.com/library/assignment-hyperlinksubaddress-property-project%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/assignment-index-property-project%28Office.15%29.aspx)|
|[LevelingDelay](http://msdn.microsoft.com/library/assignment-levelingdelay-property-project%28Office.15%29.aspx)|
|[LinkedFields](http://msdn.microsoft.com/library/assignment-linkedfields-property-project%28Office.15%29.aspx)|
|[Notes](http://msdn.microsoft.com/library/assignment-notes-property-project%28Office.15%29.aspx)|
|[Number1](http://msdn.microsoft.com/library/assignment-number1-property-project%28Office.15%29.aspx)|
|[Number10](http://msdn.microsoft.com/library/assignment-number10-property-project%28Office.15%29.aspx)|
|[Number11](http://msdn.microsoft.com/library/assignment-number11-property-project%28Office.15%29.aspx)|
|[Number12](http://msdn.microsoft.com/library/assignment-number12-property-project%28Office.15%29.aspx)|
|[Number13](http://msdn.microsoft.com/library/assignment-number13-property-project%28Office.15%29.aspx)|
|[Number14](http://msdn.microsoft.com/library/assignment-number14-property-project%28Office.15%29.aspx)|
|[Number15](http://msdn.microsoft.com/library/assignment-number15-property-project%28Office.15%29.aspx)|
|[Number16](http://msdn.microsoft.com/library/assignment-number16-property-project%28Office.15%29.aspx)|
|[Number17](http://msdn.microsoft.com/library/assignment-number17-property-project%28Office.15%29.aspx)|
|[Number18](http://msdn.microsoft.com/library/assignment-number18-property-project%28Office.15%29.aspx)|
|[Number19](http://msdn.microsoft.com/library/assignment-number19-property-project%28Office.15%29.aspx)|
|[Number2](http://msdn.microsoft.com/library/assignment-number2-property-project%28Office.15%29.aspx)|
|[Number20](http://msdn.microsoft.com/library/assignment-number20-property-project%28Office.15%29.aspx)|
|[Number3](http://msdn.microsoft.com/library/assignment-number3-property-project%28Office.15%29.aspx)|
|[Number4](http://msdn.microsoft.com/library/assignment-number4-property-project%28Office.15%29.aspx)|
|[Number5](http://msdn.microsoft.com/library/assignment-number5-property-project%28Office.15%29.aspx)|
|[Number6](http://msdn.microsoft.com/library/assignment-number6-property-project%28Office.15%29.aspx)|
|[Number7](http://msdn.microsoft.com/library/assignment-number7-property-project%28Office.15%29.aspx)|
|[Number8](http://msdn.microsoft.com/library/assignment-number8-property-project%28Office.15%29.aspx)|
|[Number9](http://msdn.microsoft.com/library/assignment-number9-property-project%28Office.15%29.aspx)|
|[Overallocated](http://msdn.microsoft.com/library/assignment-overallocated-property-project%28Office.15%29.aspx)|
|[OvertimeCost](http://msdn.microsoft.com/library/assignment-overtimecost-property-project%28Office.15%29.aspx)|
|[OvertimeWork](http://msdn.microsoft.com/library/assignment-overtimework-property-project%28Office.15%29.aspx)|
|[Owner](http://msdn.microsoft.com/library/assignment-owner-property-project%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/assignment-parent-property-project%28Office.15%29.aspx)|
|[Peak](http://msdn.microsoft.com/library/assignment-peak-property-project%28Office.15%29.aspx)|
|[PercentWorkComplete](http://msdn.microsoft.com/library/assignment-percentworkcomplete-property-project%28Office.15%29.aspx)|
|[Project](http://msdn.microsoft.com/library/assignment-project-property-project%28Office.15%29.aspx)|
|[RegularWork](http://msdn.microsoft.com/library/assignment-regularwork-property-project%28Office.15%29.aspx)|
|[RemainingCost](http://msdn.microsoft.com/library/assignment-remainingcost-property-project%28Office.15%29.aspx)|
|[RemainingOvertimeCost](http://msdn.microsoft.com/library/assignment-remainingovertimecost-property-project%28Office.15%29.aspx)|
|[RemainingOvertimeWork](http://msdn.microsoft.com/library/assignment-remainingovertimework-property-project%28Office.15%29.aspx)|
|[RemainingWork](http://msdn.microsoft.com/library/assignment-remainingwork-property-project%28Office.15%29.aspx)|
|[Resource](http://msdn.microsoft.com/library/assignment-resource-property-project%28Office.15%29.aspx)|
|[ResourceGuid](http://msdn.microsoft.com/library/assignment-resourceguid-property-project%28Office.15%29.aspx)|
|[ResourceID](http://msdn.microsoft.com/library/assignment-resourceid-property-project%28Office.15%29.aspx)|
|[ResourceName](http://msdn.microsoft.com/library/assignment-resourcename-property-project%28Office.15%29.aspx)|
|[ResourceRequestType](http://msdn.microsoft.com/library/assignment-resourcerequesttype-property-project%28Office.15%29.aspx)|
|[ResourceType](http://msdn.microsoft.com/library/assignment-resourcetype-property-project%28Office.15%29.aspx)|
|[ResourceUniqueID](http://msdn.microsoft.com/library/assignment-resourceuniqueid-property-project%28Office.15%29.aspx)|
|[ResponsePending](http://msdn.microsoft.com/library/assignment-responsepending-property-project%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/assignment-start-property-project%28Office.15%29.aspx)|
|[Start1](http://msdn.microsoft.com/library/assignment-start1-property-project%28Office.15%29.aspx)|
|[Start10](http://msdn.microsoft.com/library/assignment-start10-property-project%28Office.15%29.aspx)|
|[Start2](http://msdn.microsoft.com/library/assignment-start2-property-project%28Office.15%29.aspx)|
|[Start3](http://msdn.microsoft.com/library/assignment-start3-property-project%28Office.15%29.aspx)|
|[Start4](http://msdn.microsoft.com/library/assignment-start4-property-project%28Office.15%29.aspx)|
|[Start5](http://msdn.microsoft.com/library/assignment-start5-property-project%28Office.15%29.aspx)|
|[Start6](http://msdn.microsoft.com/library/assignment-start6-property-project%28Office.15%29.aspx)|
|[Start7](http://msdn.microsoft.com/library/assignment-start7-property-project%28Office.15%29.aspx)|
|[Start8](http://msdn.microsoft.com/library/assignment-start8-property-project%28Office.15%29.aspx)|
|[Start9](http://msdn.microsoft.com/library/assignment-start9-property-project%28Office.15%29.aspx)|
|[StartVariance](http://msdn.microsoft.com/library/assignment-startvariance-property-project%28Office.15%29.aspx)|
|[Summary](http://msdn.microsoft.com/library/assignment-summary-property-project%28Office.15%29.aspx)|
|[SV](http://msdn.microsoft.com/library/assignment-sv-property-project%28Office.15%29.aspx)|
|[Task](http://msdn.microsoft.com/library/assignment-task-property-project%28Office.15%29.aspx)|
|[TaskGuid](http://msdn.microsoft.com/library/assignment-taskguid-property-project%28Office.15%29.aspx)|
|[TaskID](http://msdn.microsoft.com/library/assignment-taskid-property-project%28Office.15%29.aspx)|
|[TaskName](http://msdn.microsoft.com/library/assignment-taskname-property-project%28Office.15%29.aspx)|
|[TaskOutlineNumber](http://msdn.microsoft.com/library/assignment-taskoutlinenumber-property-project%28Office.15%29.aspx)|
|[TaskSummaryName](http://msdn.microsoft.com/library/assignment-tasksummaryname-property-project%28Office.15%29.aspx)|
|[TaskUniqueID](http://msdn.microsoft.com/library/assignment-taskuniqueid-property-project%28Office.15%29.aspx)|
|[TeamStatusPending](http://msdn.microsoft.com/library/assignment-teamstatuspending-property-project%28Office.15%29.aspx)|
|[Text1](http://msdn.microsoft.com/library/assignment-text1-property-project%28Office.15%29.aspx)|
|[Text10](http://msdn.microsoft.com/library/assignment-text10-property-project%28Office.15%29.aspx)|
|[Text11](http://msdn.microsoft.com/library/assignment-text11-property-project%28Office.15%29.aspx)|
|[Text12](http://msdn.microsoft.com/library/assignment-text12-property-project%28Office.15%29.aspx)|
|[Text13](http://msdn.microsoft.com/library/assignment-text13-property-project%28Office.15%29.aspx)|
|[Text14](http://msdn.microsoft.com/library/assignment-text14-property-project%28Office.15%29.aspx)|
|[Text15](http://msdn.microsoft.com/library/assignment-text15-property-project%28Office.15%29.aspx)|
|[Text16](http://msdn.microsoft.com/library/assignment-text16-property-project%28Office.15%29.aspx)|
|[Text17](http://msdn.microsoft.com/library/assignment-text17-property-project%28Office.15%29.aspx)|
|[Text18](http://msdn.microsoft.com/library/assignment-text18-property-project%28Office.15%29.aspx)|
|[Text19](http://msdn.microsoft.com/library/assignment-text19-property-project%28Office.15%29.aspx)|
|[Text2](http://msdn.microsoft.com/library/assignment-text2-property-project%28Office.15%29.aspx)|
|[Text20](http://msdn.microsoft.com/library/assignment-text20-property-project%28Office.15%29.aspx)|
|[Text21](http://msdn.microsoft.com/library/assignment-text21-property-project%28Office.15%29.aspx)|
|[Text22](http://msdn.microsoft.com/library/assignment-text22-property-project%28Office.15%29.aspx)|
|[Text23](http://msdn.microsoft.com/library/assignment-text23-property-project%28Office.15%29.aspx)|
|[Text24](http://msdn.microsoft.com/library/assignment-text24-property-project%28Office.15%29.aspx)|
|[Text25](http://msdn.microsoft.com/library/assignment-text25-property-project%28Office.15%29.aspx)|
|[Text26](http://msdn.microsoft.com/library/assignment-text26-property-project%28Office.15%29.aspx)|
|[Text27](http://msdn.microsoft.com/library/assignment-text27-property-project%28Office.15%29.aspx)|
|[Text28](http://msdn.microsoft.com/library/assignment-text28-property-project%28Office.15%29.aspx)|
|[Text29](http://msdn.microsoft.com/library/assignment-text29-property-project%28Office.15%29.aspx)|
|[Text3](http://msdn.microsoft.com/library/assignment-text3-property-project%28Office.15%29.aspx)|
|[Text30](http://msdn.microsoft.com/library/assignment-text30-property-project%28Office.15%29.aspx)|
|[Text4](http://msdn.microsoft.com/library/assignment-text4-property-project%28Office.15%29.aspx)|
|[Text5](http://msdn.microsoft.com/library/assignment-text5-property-project%28Office.15%29.aspx)|
|[Text6](http://msdn.microsoft.com/library/assignment-text6-property-project%28Office.15%29.aspx)|
|[Text7](http://msdn.microsoft.com/library/assignment-text7-property-project%28Office.15%29.aspx)|
|[Text8](http://msdn.microsoft.com/library/assignment-text8-property-project%28Office.15%29.aspx)|
|[Text9](http://msdn.microsoft.com/library/assignment-text9-property-project%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/assignment-uniqueid-property-project%28Office.15%29.aspx)|
|[Units](http://msdn.microsoft.com/library/assignment-units-property-project%28Office.15%29.aspx)|
|[UpdateNeeded](http://msdn.microsoft.com/library/assignment-updateneeded-property-project%28Office.15%29.aspx)|
|[VAC](http://msdn.microsoft.com/library/assignment-vac-property-project%28Office.15%29.aspx)|
|[WBS](http://msdn.microsoft.com/library/assignment-wbs-property-project%28Office.15%29.aspx)|
|[Work](http://msdn.microsoft.com/library/assignment-work-property-project%28Office.15%29.aspx)|
|[WorkContour](http://msdn.microsoft.com/library/assignment-workcontour-property-project%28Office.15%29.aspx)|
|[WorkVariance](http://msdn.microsoft.com/library/assignment-workvariance-property-project%28Office.15%29.aspx)|
|[Compliant](http://msdn.microsoft.com/library/assignment-compliant-property-project%28Office.15%29.aspx)|

