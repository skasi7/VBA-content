---
title: Task Members (Project)
ms.prod: PROJECTSERVER
ms.assetid: abbe80c2-4458-5c3a-5b9c-095759c9fce4
---


# Task Members (Project)





## Methods



|**Name**|**Description**|
|:-----|:-----|
|[AppendNotes](task-appendnotes-method-project.md)|Appends text to the Notes field.|
|[Delete](task-delete-method-project.md)|Deletes the  **Task** object from a **Tasks** collection.|
|[GetField](task-getfield-method-project.md)|Returns the value of the specified task custom field.|
|[LinkPredecessors](task-linkpredecessors-method-project.md)|Adds one or more predecessors to the task.|
|[LinkSuccessors](task-linksuccessors-method-project.md)|Adds one or more successors to the task.|
|[OutlineHideSubTasks](task-outlinehidesubtasks-method-project.md)|Hides the subtasks of the selected task or tasks.|
|[OutlineIndent](task-outlineindent-method-project.md)|Indents a task in the outline.|
|[OutlineOutdent](task-outlineoutdent-method-project.md)|Promotes a task in the outline.|
|[OutlineShowAllTasks](task-outlineshowalltasks-method-project.md)|Expands all summary tasks in the project.|
|[OutlineShowSubTasks](task-outlineshowsubtasks-method-project.md)|Shows the subtasks of the selected task or tasks.|
|[SetField](task-setfield-method-project.md)|Sets the value of the specified task custom field.|
|[Split](task-split-method-project.md)|Splits the task into two portions.|
|[TimeScaleData](task-timescaledata-method-project.md)|Sets options for displaying timephased data for the task.|
|[UnlinkPredecessors](task-unlinkpredecessors-method-project.md)|Removes one or more predecessors from the task.|
|[UnlinkSuccessors](task-unlinksuccessors-method-project.md)|Removes one or more successors from the task.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Active](task-active-property-project.md)|**True** if the task is active; otherwise, **False**. Read/write **Variant**.|
|[ActualCost](task-actualcost-property-project.md)|Gets or sets the actual cost for the task. Read/write  **Variant**.|
|[ActualDuration](task-actualduration-property-project.md)|Gets or sets the actual duration (in minutes) of a task. Read-only for summary tasks. Read/write  **Variant**.|
|[ActualFinish](task-actualfinish-property-project.md)|Gets or sets the actual finish date of a task. Read-only for summary tasks. Read/write  **Variant**.|
|[ActualOvertimeCost](task-actualovertimecost-property-project.md)|Gets the actual overtime cost for a task. Read-only  **Variant**.|
|[ActualOvertimeWork](task-actualovertimework-property-project.md)|Gets the actual overtime work (in minutes) for a task. Read-only  **Variant**.|
|[ActualStart](task-actualstart-property-project.md)|Gets or sets the actual start date of the task. Read-only for summary tasks. Read/write  **Variant**.|
|[ActualWork](task-actualwork-property-project.md)|Gets or sets the actual work (in minutes) for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[ACWP](task-acwp-property-project.md)|Gets the actual cost of work performed for the task. Read-only  **Variant**.|
|[Application](task-application-property-project.md)|Gets the  **[Application](application-object-project.md)** object. Read-only **Application**.|
|[Assignments](task-assignments-property-project.md)|Gets an  **[Assignments](assignment-object-project.md)** collection representing the assignments for the task. Read-only **Assignments**.|
|[Baseline10BudgetCost](task-baseline10budgetcost-property-project.md)|Gets or sets the baseline10 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline10BudgetWork](task-baseline10budgetwork-property-project.md)|Gets or sets the baseline10 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline10Cost](task-baseline10cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline10DeliverableFinish](task-baseline10deliverablefinish-property-project.md)|Gets or sets the task baseline10 deliverables finish date. Read/write  **Variant**.|
|[Baseline10DeliverableStart](task-baseline10deliverablestart-property-project.md)|Gets or sets the task baseline10 deliverables start date. Read/write  **Variant**.|
|[Baseline10Duration](task-baseline10duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline10DurationEstimated](task-baseline10durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline10DurationText](task-baseline10durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline10Finish](task-baseline10finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline10FinishText](task-baseline10finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline10FixedCost](task-baseline10fixedcost-property-project.md)|Gets or sets the baseline10 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline10FixedCostAccrual](task-baseline10fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline10 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline10Start](task-baseline10start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline10StartText](task-baseline10starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline10Work](task-baseline10work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline1BudgetCost](task-baseline1budgetcost-property-project.md)|Gets or sets the baseline1 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline1BudgetWork](task-baseline1budgetwork-property-project.md)|Gets or sets the baseline1 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline1Cost](task-baseline1cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline1DeliverableFinish](task-baseline1deliverablefinish-property-project.md)|Gets or sets the task baseline1 deliverables finish date. Read/write  **Variant**.|
|[Baseline1DeliverableStart](task-baseline1deliverablestart-property-project.md)|Gets or sets the task baseline1 deliverables start date. Read/write  **Variant**.|
|[Baseline1Duration](task-baseline1duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline1DurationEstimated](task-baseline1durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline1DurationText](task-baseline1durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline1Finish](task-baseline1finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline1FinishText](task-baseline1finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline1FixedCost](task-baseline1fixedcost-property-project.md)|Gets or sets the baseline1 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline1FixedCostAccrual](task-baseline1fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline1 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline1Start](task-baseline1start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline1StartText](task-baseline1starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline1Work](task-baseline1work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline2BudgetCost](task-baseline2budgetcost-property-project.md)|Gets or sets the baseline2 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline2BudgetWork](task-baseline2budgetwork-property-project.md)|Gets or sets the baseline2 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline2Cost](task-baseline2cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline2DeliverableFinish](task-baseline2deliverablefinish-property-project.md)|Gets or sets the task baseline2 deliverables finish date. Read/write  **Variant**.|
|[Baseline2DeliverableStart](task-baseline2deliverablestart-property-project.md)|Gets or sets the task baseline2 deliverables start date. Read/write  **Variant**.|
|[Baseline2Duration](task-baseline2duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline2DurationEstimated](task-baseline2durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline2DurationText](task-baseline2durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline2Finish](task-baseline2finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline2FinishText](task-baseline2finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline2FixedCost](task-baseline2fixedcost-property-project.md)|Gets or sets the baseline2 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline2FixedCostAccrual](task-baseline2fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline2 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline2Start](task-baseline2start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline2StartText](task-baseline2starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline2Work](task-baseline2work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline3BudgetCost](task-baseline3budgetcost-property-project.md)|Gets or sets the baseline3 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline3BudgetWork](task-baseline3budgetwork-property-project.md)|Gets or sets the baseline3 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline3Cost](task-baseline3cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline3DeliverableFinish](task-baseline3deliverablefinish-property-project.md)|Gets or sets the task baseline3 deliverables finish date. Read/write  **Variant**.|
|[Baseline3DeliverableStart](task-baseline3deliverablestart-property-project.md)|Gets or sets the task baseline3 deliverables start date. Read/write  **Variant**.|
|[Baseline3Duration](task-baseline3duration-property-project.md)|Gets or sets the baseline3 duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline3DurationEstimated](task-baseline3durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline3DurationText](task-baseline3durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline3Finish](task-baseline3finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline3FinishText](task-baseline3finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline3FixedCost](task-baseline3fixedcost-property-project.md)|Gets or sets the baseline3 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline3FixedCostAccrual](task-baseline3fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline3 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline3Start](task-baseline3start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline3StartText](task-baseline3starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline3Work](task-baseline3work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline4BudgetCost](task-baseline4budgetcost-property-project.md)|Gets or sets the baseline4 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline4BudgetWork](task-baseline4budgetwork-property-project.md)|Gets or sets the baseline4 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline4Cost](task-baseline4cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline4DeliverableFinish](task-baseline4deliverablefinish-property-project.md)|Gets or sets the task baseline4 deliverables finish date. Read/write  **Variant**.|
|[Baseline4DeliverableStart](task-baseline4deliverablestart-property-project.md)|Gets or sets the task baseline4 deliverables start date. Read/write  **Variant**.|
|[Baseline4Duration](task-baseline4duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline4DurationEstimated](task-baseline4durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline4DurationText](task-baseline4durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline4Finish](task-baseline4finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline4FinishText](task-baseline4finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline4FixedCost](task-baseline4fixedcost-property-project.md)|Gets or sets the baseline4 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline4FixedCostAccrual](task-baseline4fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline4 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline4Start](task-baseline4start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline4StartText](task-baseline4starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline4Work](task-baseline4work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline5BudgetCost](task-baseline5budgetcost-property-project.md)|Gets or sets the baseline5 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline5BudgetWork](task-baseline5budgetwork-property-project.md)|Gets or sets the baseline5 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline5Cost](task-baseline5cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline5DeliverableFinish](task-baseline5deliverablefinish-property-project.md)|Gets or sets the task baseline5 deliverables finish date. Read/write  **Variant**.|
|[Baseline5DeliverableStart](task-baseline5deliverablestart-property-project.md)|Gets or sets the task baseline5 deliverables start date. Read/write  **Variant**.|
|[Baseline5Duration](task-baseline5duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline5DurationEstimated](task-baseline5durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline5DurationText](task-baseline5durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline5Finish](task-baseline5finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline5FinishText](task-baseline5finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline5FixedCost](task-baseline5fixedcost-property-project.md)|Gets or sets the baseline5 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline5FixedCostAccrual](task-baseline5fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline5 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline5Start](task-baseline5start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline5StartText](task-baseline5starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline5Work](task-baseline5work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline6BudgetCost](task-baseline6budgetcost-property-project.md)|Gets or sets the baseline6 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline6BudgetWork](task-baseline6budgetwork-property-project.md)|Gets or sets the baseline6 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline6Cost](task-baseline6cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline6DeliverableFinish](task-baseline6deliverablefinish-property-project.md)|Gets or sets the task baseline6 deliverables finish date. Read/write  **Variant**.|
|[Baseline6DeliverableStart](task-baseline6deliverablestart-property-project.md)|Gets or sets the task baseline6 deliverables start date. Read/write  **Variant**.|
|[Baseline6Duration](task-baseline6duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline6DurationEstimated](task-baseline6durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline6DurationText](task-baseline6durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline6Finish](task-baseline6finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline6FinishText](task-baseline6finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline6FixedCost](task-baseline6fixedcost-property-project.md)|Gets or sets the baseline6 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline6FixedCostAccrual](task-baseline6fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline6 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline6Start](task-baseline6start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline6StartText](task-baseline6starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline6Work](task-baseline6work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline7BudgetCost](task-baseline7budgetcost-property-project.md)|Gets or sets the baseline7 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline7BudgetWork](task-baseline7budgetwork-property-project.md)|Gets or sets the baseline7 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline7Cost](task-baseline7cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline7DeliverableFinish](task-baseline7deliverablefinish-property-project.md)|Gets or sets the task baseline7 deliverables finish date. Read/write  **Variant**.|
|[Baseline7DeliverableStart](task-baseline7deliverablestart-property-project.md)|Gets or sets the task baseline7 deliverables start date. Read/write  **Variant**.|
|[Baseline7Duration](task-baseline7duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline7DurationEstimated](task-baseline7durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline7DurationText](task-baseline7durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline7Finish](task-baseline7finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline7FinishText](task-baseline7finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline7FixedCost](task-baseline7fixedcost-property-project.md)|Gets or sets the baseline7 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline7FixedCostAccrual](task-baseline7fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline7 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline7Start](task-baseline7start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline7StartText](task-baseline7starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline7Work](task-baseline7work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline8BudgetCost](task-baseline8budgetcost-property-project.md)|Gets or sets the baseline8 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline8BudgetWork](task-baseline8budgetwork-property-project.md)|Gets or sets the baseline8 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline8Cost](task-baseline8cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline8DeliverableFinish](task-baseline8deliverablefinish-property-project.md)|Gets or sets the task baseline8 deliverables finish date. Read/write  **Variant**.|
|[Baseline8DeliverableStart](task-baseline8deliverablestart-property-project.md)|Gets or sets the task baseline8 deliverables start date. Read/write  **Variant**.|
|[Baseline8Duration](task-baseline8duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline8DurationEstimated](task-baseline8durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline8DurationText](task-baseline8durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline8Finish](task-baseline8finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline8FinishText](task-baseline8finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline8FixedCost](task-baseline8fixedcost-property-project.md)|Gets or sets the baseline8 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline8FixedCostAccrual](task-baseline8fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline8 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline8Start](task-baseline8start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline8StartText](task-baseline8starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline8Work](task-baseline8work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[Baseline9BudgetCost](task-baseline9budgetcost-property-project.md)|Gets or sets the baseline9 budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline9BudgetWork](task-baseline9budgetwork-property-project.md)|Gets or sets the baseline9 budget work for the rollup calculated budgeted work hours for all the work and the material resources for the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Baseline9Cost](task-baseline9cost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[Baseline9DeliverableFinish](task-baseline9deliverablefinish-property-project.md)|Gets or sets the task baseline9 deliverables finish date. Read/write  **Variant**.|
|[Baseline9DeliverableStart](task-baseline9deliverablestart-property-project.md)|Gets or sets the task baseline9 deliverables start date. Read/write  **Variant**.|
|[Baseline9Duration](task-baseline9duration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[Baseline9DurationEstimated](task-baseline9durationestimated-property-project.md)|**True** if the baseline duration of a task is an estimate. Read/write **Variant**.|
|[Baseline9DurationText](task-baseline9durationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[Baseline9Finish](task-baseline9finish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[Baseline9FinishText](task-baseline9finishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[Baseline9FixedCost](task-baseline9fixedcost-property-project.md)|Gets or sets the baseline9 fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[Baseline9FixedCostAccrual](task-baseline9fixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline9 accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[Baseline9Start](task-baseline9start-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[Baseline9StartText](task-baseline9starttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[Baseline9Work](task-baseline9work-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[BaselineBudgetCost](task-baselinebudgetcost-property-project.md)|Gets or sets the baseline budget cost for the rollup calculated value of all the cost resources within the project. Applies only to the project summary task. Read/write  **Variant**.|
|[BaselineBudgetWork](task-baselinebudgetwork-property-project.md)|Gets or sets the baseline budget work hours for all non-cost resources assigned to the project summary task. Read/write  **Variant**.|
|[BaselineCost](task-baselinecost-property-project.md)|Gets or sets the baseline cost for a  **Task**. Read/write **Variant**.|
|[BaselineDeliverableFinish](task-baselinedeliverablefinish-property-project.md)|Gets or sets the task baseline deliverables finish date. Read/write  **Variant**.|
|[BaselineDeliverableStart](task-baselinedeliverablestart-property-project.md)|Gets or sets the task baseline deliverables start date. Read/write  **Variant**.|
|[BaselineDuration](task-baselineduration-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[BaselineDurationEstimated](task-baselinedurationestimated-property-project.md)|Gets or sets the baseline duration (in minutes) of a task. Read/write  **Variant**.|
|[BaselineDurationText](task-baselinedurationtext-property-project.md)|Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.|
|[BaselineFinish](task-baselinefinish-property-project.md)|Gets or sets the baseline finish date of a  **Task**. Read/write **Variant**.|
|[BaselineFinishText](task-baselinefinishtext-property-project.md)|Gets or sets a string representation of the baseline finish date of a task. Read/write  **String**.|
|[BaselineFixedCost](task-baselinefixedcost-property-project.md)|Gets or sets the baseline fixed cost of any nonresource expense for a  **Task**. Read/write **Variant**.|
|[BaselineFixedCostAccrual](task-baselinefixedcostaccrual-property-project.md)|Gets or sets when the  **Task** baseline accrues fixed costs. Read/write **Long**. Can be one of the **[PjAccrueAt](pjaccrueat-enumeration-project.md)** constants.|
|[BaselineStart](task-baselinestart-property-project.md)|Gets or sets the baseline start date of a  **Task**. Read/write **Variant**.|
|[BaselineStartText](task-baselinestarttext-property-project.md)|Gets or sets a string representation of the baseline start date of a task. Read/write  **String**.|
|[BaselineWork](task-baselinework-property-project.md)|Gets or sets the baseline work (in minutes) for a  **Task**. Read/write **Variant**.|
|[BCWP](task-bcwp-property-project.md)|Gets the budgeted cost of work performed for the task. Read-only  **Variant**.|
|[BCWS](task-bcws-property-project.md)|Gets the budgeted cost of work scheduled for the task. Read-only  **Variant**.|
|[BudgetCost](task-budgetcost-property-project.md)| Gets or sets the rollup calculated value of the budget costs for all the cost resources within the project. Applies only to the project summary task. Read/write **Variant**.|
|[BudgetWork](task-budgetwork-property-project.md)|Gets or sets the rollup calculated value of all the budget work hours across all the non-cost resources in the project. Applies only to the project summary task. Read/write  **Variant**.|
|[Calendar](task-calendar-property-project.md)|Gets or sets the name of the calendar to be used when scheduling the task. Read/write  **String**.|
|[CalendarGuid](task-calendarguid-property-project.md)|Gets the GUID of the calendar for the task. Read-only  **String**.|
|[CalendarObject](task-calendarobject-property-project.md)|Gets the Calendar object to be used when scheduling the task. Read-only  **Calendar**.|
|[Confirmed](task-confirmed-property-project.md)|Gets the results of task assignments in a Project mail message.  **True** if all resources assigned to the task have accepted their assignments. Read-only **Boolean**.|
|[ConstraintDate](task-constraintdate-property-project.md)|Gets or sets a constraint date for a task. Read/write  **Variant**.|
|[ConstraintType](task-constrainttype-property-project.md)|Gets or sets a constraint type for a task. Read/write  **Variant**.|
|[Contact](task-contact-property-project.md)|Gets or sets contact information of the person who is responsible for a task. Read/write  **String**.|
|[Cost](task-cost-property-project.md)|Gets the total cost of the task. Read/write  **Variant**.|
|[Cost1](task-cost1-property-project.md)|Gets or sets the value of the  **Cost1** custom field for the task. Read/write **Variant**.|
|[Cost10](task-cost10-property-project.md)|Gets or sets the value of the  **Cost10** custom field for the task. Read/write **Variant**.|
|[Cost2](task-cost2-property-project.md)|Gets or sets the value of the  **Cost2** custom field for the task. Read/write **Variant**.|
|[Cost3](task-cost3-property-project.md)|Gets or sets the value of the  **Cost3** custom field for the task. Read/write **Variant**.|
|[Cost4](task-cost4-property-project.md)|Gets or sets the value of the  **Cost4** custom field for the task. Read/write **Variant**.|
|[Cost5](task-cost5-property-project.md)|Gets or sets the value of the  **Cost5** custom field for the task. Read/write **Variant**.|
|[Cost6](task-cost6-property-project.md)|Gets or sets the value of the  **Cost6** custom field for the task. Read/write **Variant**.|
|[Cost7](task-cost7-property-project.md)|Gets or sets the value of the  **Cost7** custom field for the task. Read/write **Variant**.|
|[Cost8](task-cost8-property-project.md)|Gets or sets the  **Cost8** custom field for the task. Read/write **Variant**.|
|[Cost9](task-cost9-property-project.md)|Gets or sets the value of the  **Cost9** custom field for the task. Read/write **Variant**.|
|[CostVariance](task-costvariance-property-project.md)|Gets the variance between the baseline cost and the cost of a  **Task**. Read-only **Variant**.|
|[CPI](task-cpi-property-project.md)|Gets the cost performance index of a specified task. Read-only  **Double**.|
|[Created](task-created-property-project.md)|Gets the date a  **Task** was created. Read-only **Variant**.|
|[Critical](task-critical-property-project.md)|**True** if the task is on the critical path. Read-only **Boolean**.|
|[CV](task-cv-property-project.md)|Gets the cost variance for a  **Task**. Read-only **Variant**.|
|[CVPercent](task-cvpercent-property-project.md)|Gets the cost variance percent of the task. Read-only  **Variant**.|
|[Date1](task-date1-property-project.md)|Gets or sets the value of the  **Date1** custom field for the task. Read/write **Variant**.|
|[Date10](task-date10-property-project.md)|Gets or sets the value of the  **Date10** custom field for the task. Read/write **Variant**.|
|[Date2](task-date2-property-project.md)|Gets or sets the value of the  **Date2** custom field for the task. Read/write **Variant**.|
|[Date3](task-date3-property-project.md)|Gets or sets the value of the  **Date3** custom field for the task. Read/write **Variant**.|
|[Date4](task-date4-property-project.md)|Gets or sets the value of the  **Date4** custom field for the task. Read/write **Variant**.|
|[Date5](task-date5-property-project.md)|Gets or sets the value of the  **Date5** custom field for the task. Read/write **Variant**.|
|[Date6](task-date6-property-project.md)|Gets or sets the value of the  **Date6** custom field for the task. Read/write **Variant**.|
|[Date7](task-date7-property-project.md)|Gets or sets the value of the  **Date7** custom field for the task. Read/write **Variant**.|
|[Date8](task-date8-property-project.md)|Gets or sets the value of the  **Date8** custom field for the task. Read/write **Variant**.|
|[Date9](task-date9-property-project.md)|Gets or sets the value of the  **Date9** custom field for the task. Read/write **Variant**.|
|[Deadline](task-deadline-property-project.md)|Gets or sets a deadline for a task. Read/write  **Variant**.|
|[DeliverableFinish](task-deliverablefinish-property-project.md)|Gets or sets the task deliverable end date. Read/write  **Variant**.|
|[DeliverableGuid](task-deliverableguid-property-project.md)|Gets or sets the GUID of the task deliverable. Read/write  **String**.|
|[DeliverableName](task-deliverablename-property-project.md)|Gets or sets the name of the deliverable. Read/write  **String**.|
|[DeliverableStart](task-deliverablestart-property-project.md)|Gets or sets the task deliverable start date. Read/write  **Variant**.|
|[DeliverableType](task-deliverabletype-property-project.md)|Gets or sets the type of deliverable for the task. Read/write  **Integer**.|
|[Duration](task-duration-property-project.md)|Gets or sets the duration (in minutes) of a task. Read-only for summary tasks. Read/write  **Variant**.|
|[Duration1](task-duration1-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration10](task-duration10-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration10Estimated](task-duration10estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration1Estimated](task-duration1estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration2](task-duration2-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration2Estimated](task-duration2estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration3](task-duration3-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration3Estimated](task-duration3estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration4](task-duration4-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration4Estimated](task-duration4estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration5](task-duration5-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration5Estimated](task-duration5estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration6](task-duration6-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration6Estimated](task-duration6estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration7](task-duration7-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration7Estimated](task-duration7estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration8](task-duration8-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration8Estimated](task-duration8estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[Duration9](task-duration9-property-project.md)| Gets or sets the value of a task duration custom field. Read/write **Variant**.|
|[Duration9Estimated](task-duration9estimated-property-project.md)|**True** if a task duration custom field is an estimate. Read/write **Variant**.|
|[DurationText](task-durationtext-property-project.md)|Gets or sets a string representation of the task duration. Read/write  **String**.|
|[DurationVariance](task-durationvariance-property-project.md)|Gets the variance (in minutes) between the planned duration and the duration of the task. Read-only  **Variant**.|
|[EAC](task-eac-property-project.md)|Gets the estimate at completion (EAC) for the task. Read-only  **Variant**.|
|[EarlyFinish](task-earlyfinish-property-project.md)|Gets the earliest date on which a task can finish. Read-only  **Variant**.|
|[EarlyStart](task-earlystart-property-project.md)|Gets the earliest date on which a task can start. Read-only  **Variant**.|
|[EarnedValueMethod](task-earnedvaluemethod-property-project.md)|Gets or sets the method for calculating earned value for a task. Read/write  **PjEarnedValueMethod**.|
|[EffortDriven](task-effortdriven-property-project.md)|**True** if the task is effort-driven. Read/write **Variant**.|
|[ErrorMessage](task-errormessage-property-project.md)|Gets errors reported by the  **Import Task Wizard** relating to custom fields and calendar validations. Read-only **String**.|
|[Estimated](task-estimated-property-project.md)|**True** if the the task duration is an estimate. **False** if the task duration is a set value. Read/write **Variant**.|
|[ExternalTask](task-externaltask-property-project.md)|**True** if the task is actually a placeholder for a task in another project. Read-only **Variant**.|
|[Finish](task-finish-property-project.md)|Gets or sets the finish date of a  **Task**. Read-only for summary tasks. Read/write **Variant**.|
|[Finish1](task-finish1-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish10](task-finish10-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish2](task-finish2-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish3](task-finish3-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish4](task-finish4-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish5](task-finish5-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish6](task-finish6-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish7](task-finish7-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish8](task-finish8-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[Finish9](task-finish9-property-project.md)|Gets or sets the local Finish custom field of the task. Read/write  **Variant**.|
|[FinishSlack](task-finishslack-property-project.md)|Gets or sets the finish slack of a task in minutes. Read-only  **Variant**.|
|[FinishText](task-finishtext-property-project.md)|Gets or sets a string representation of the task finish date. Read/write  **String**.|
|[FinishVariance](task-finishvariance-property-project.md)|Gets the variance (in minutes) between the baseline finish date and the finish date of a task. Read-only  **Variant**.|
|[FixedCost](task-fixedcost-property-project.md)|Gets or sets a fixed cost for a task. Read/write  **Variant**.|
|[FixedCostAccrual](task-fixedcostaccrual-property-project.md)|Gets or sets the way the task accrues fixed costs. Read/write  **PjAccrueAt**.|
|[Flag1](task-flag1-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag10](task-flag10-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag11](task-flag11-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag12](task-flag12-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag13](task-flag13-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag14](task-flag14-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag15](task-flag15-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag16](task-flag16-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag17](task-flag17-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag18](task-flag18-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag19](task-flag19-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag2](task-flag2-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag20](task-flag20-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag3](task-flag3-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag4](task-flag4-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag5](task-flag5-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag6](task-flag6-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag7](task-flag7-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag8](task-flag8-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[Flag9](task-flag9-property-project.md)|Gets or sets the value of a task flag custom field. Read/write  **Variant**.|
|[FreeSlack](task-freeslack-property-project.md)|Gets the free slack for a task in minutes. Read-only  **Variant**.|
|[GroupBySummary](task-groupbysummary-property-project.md)|**True** if the selected item in a task view is in a group summary row; otherwise, **false**. Read-only **Boolean**.|
|[Guid](task-guid-property-project.md)|Gets the GUID of the task. Read-only  **String**.|
|[HideBar](task-hidebar-property-project.md)|**True** if a task bar does not appear on the Gantt Chart or Calendar. Read/write **Variant**.|
|[Hyperlink](task-hyperlink-property-project.md)|Gets or sets a friendly name representing a hyperlink address. The name may also be a URL or UNC path. Read/write  **String**.|
|[HyperlinkAddress](task-hyperlinkaddress-property-project.md)|Gets or sets the URL or UNC path of a document. Read/write  **String**.|
|[HyperlinkHREF](task-hyperlinkhref-property-project.md)|Gets or sets a combination of the hyperlink address and subaddress, separated by a "#". Read/write  **String**.|
|[HyperlinkScreenTip](task-hyperlinkscreentip-property-project.md)|Gets or sets a ScreenTip for the hyperlink. Read/write  **String**.|
|[HyperlinkSubAddress](task-hyperlinksubaddress-property-project.md)|Gets or sets the address of a location within the target document. Read/write  **String**.|
|[ID](task-id-property-project.md)|Gets the identification number of a task. Read-only  **Long**.|
|[IgnoreResourceCalendar](task-ignoreresourcecalendar-property-project.md)|**True** if the resource calendar is ignored when scheduling the task. **False** if both the resource calendar and task calendar (if defined) are used when scheduling the task. Read/write **Variant**.|
|[IgnoreWarnings](task-ignorewarnings-property-project.md)|**True** if task warnings are ignored when processing the task; otherwise, **False**. Read/write **Variant**.|
|[Index](task-index-property-project.md)|Gets the index of a  **Task** object in the **Tasks** containing object. Read-only **Long**.|
|[IsDurationValid](task-isdurationvalid-property-project.md)|**True** if the duration of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.|
|[IsFinishValid](task-isfinishvalid-property-project.md)|**True** if the finish date of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.|
|[IsPublished](task-ispublished-property-project.md)|**True** when the task and its assignments are published. Read/write **Variant**.|
|[IsStartValid](task-isstartvalid-property-project.md)|**True** if the start date of a manually scheduled task is valid; otherwise, **False**. Read-only **Boolean**.|
|[LateFinish](task-latefinish-property-project.md)|Gets the latest date on which a task can finish. Read-only  **Variant**.|
|[LateStart](task-latestart-property-project.md)|Gets the latest date on which a task can start. Read-only  **Variant**.|
|[LevelIndividualAssignments](task-levelindividualassignments-property-project.md)|**True** if leveling can adjust individual assignments on a task. **False** if all assignments on a task can be adjusted, including those that are not overallocated. Read/write **Variant**.|
|[LevelingCanSplit](task-levelingcansplit-property-project.md)|**True** if leveling can create splits in remaining work. Read/write **Boolean**.|
|[LevelingDelay](task-levelingdelay-property-project.md)|Gets or sets the amount of time the task is delayed due to leveling. Read/write  **Variant**.|
|[LinkedFields](task-linkedfields-property-project.md)|**True** if the **Task** object contains fields that are linked to other applications through OLE. Read-only **Boolean**.|
|[Manual](task-manual-property-project.md)|**True** if task recalculation is set to **Manually Scheduled**;  **False** if task recalculation is set to **Auto Schedule**. Read/write  **Variant**.|
|[Marked](task-marked-property-project.md)|**True** if the task is marked for further action of some kind. Read/write **Variant**.|
|[Milestone](task-milestone-property-project.md)|**True** if the task is a milestone. Read/write **Variant**.|
|[Name](task-name-property-project.md)|Gets or sets the name of a  **Task** object. Read/write **String**.|
|[Notes](task-notes-property-project.md)|Gets or sets the notes for a task. Read/write  **String**.|
|[Number1](task-number1-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number10](task-number10-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number11](task-number11-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number12](task-number12-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number13](task-number13-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number14](task-number14-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number15](task-number15-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number16](task-number16-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number17](task-number17-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number18](task-number18-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number19](task-number19-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number2](task-number2-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number20](task-number20-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number3](task-number3-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number4](task-number4-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number5](task-number5-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number6](task-number6-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number7](task-number7-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number8](task-number8-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Number9](task-number9-property-project.md)|Gets or sets a Number local custom field for a task. Read/write  **Double**.|
|[Objects](task-objects-property-project.md)|Gets the number of OLE objects contained within a  **Task** object. Any objects inserted in the Notes field of a task are not included in the count. Read-only **Long**.|
|[OutlineChildren](task-outlinechildren-property-project.md)|Gets a  **[Tasks](task-object-project.md)** collection representing the children of a task in the outline structure. Read-only **Tasks**.|
|[OutlineCode1](task-outlinecode1-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode10](task-outlinecode10-property-project.md)| Returns or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode2](task-outlinecode2-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode3](task-outlinecode3-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode4](task-outlinecode4-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode5](task-outlinecode5-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode6](task-outlinecode6-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode7](task-outlinecode7-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode8](task-outlinecode8-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineCode9](task-outlinecode9-property-project.md)| Gets or sets the value of the outline code custom field for a task. Read/write **String**.|
|[OutlineLevel](task-outlinelevel-property-project.md)|Gets the level of the task in the outline hierarchy. Read/write  **Integer**.|
|[OutlineNumber](task-outlinenumber-property-project.md)|Gets a value that indicates the position of the task in the outline hierarchy. Read-only  **String**.|
|[OutlineParent](task-outlineparent-property-project.md)|Gets a  **[Task](task-object-project.md)** object representing the parent of a task in the outline structure. Read-only **Task**.|
|[Overallocated](task-overallocated-property-project.md)|**True** if any of the assignments for a task is overallocated. Read-only **Boolean**.|
|[OvertimeCost](task-overtimecost-property-project.md)|Gets the overtime cost for a task. Read-only  **Variant**.|
|[OvertimeWork](task-overtimework-property-project.md)|Gets the overtime work for a task. Read-only  **Variant**.|
|[Parent](task-parent-property-project.md)|Gets the parent of the  **Task** object. Read-only **Object**.|
|[PathDrivenSuccessor](task-pathdrivensuccessor-property-project.md)|Gets a value that indicates whether the task is a successor that is driven by the selected task, when the  **DrivenSuccessors** item is selected in the **Task Path** drop-down list. Read-only **Boolean**.|
|[PathDrivingPredecessor](task-pathdrivingpredecessor-property-project.md)|Gets a value that indicates whether the task is a predecessor that drives the selected task, when the  **Driving Predecessors** item is selected in the **Task Path** drop-down list. Read-only **Boolean**.|
|[PathPredecessor](task-pathpredecessor-property-project.md)|Gets a value that indicates whether the task is a predecessor of the selected task, when the  **Predecessors** item is selected in the **Task Path** drop-down list. Read-only **Boolean**.|
|[PathSuccessor](task-pathsuccessor-property-project.md)|Gets a value that indicates whether the task is a successor of the selected task, when the  **Successors** item is selected in the **Task Path** drop-down list. Read-only **Boolean**.|
|[PercentComplete](task-percentcomplete-property-project.md)|Gets or sets the percent complete of a task. Read/write  **Variant**.|
|[PercentWorkComplete](task-percentworkcomplete-property-project.md)|Gets or sets the percentage of work complete for a task. Read-only for summary tasks. Read/write  **Variant**.|
|[PhysicalPercentComplete](task-physicalpercentcomplete-property-project.md)|Gets or sets the physical percent complete of a task. Read/write  **Variant**.|
|[Placeholder](task-placeholder-property-project.md)|**True** if the task is a placeholder for another task. Read-only **Variant**.|
|[Predecessors](task-predecessors-property-project.md)|Gets or sets a list of the identification numbers of a task's predecessors. Read/write  **String**.|
|[PredecessorTasks](task-predecessortasks-property-project.md)|Gets a  **[Tasks](task-object-project.md)** collection representing the predecessors of the task. Read-only **Tasks**.|
|[PreleveledFinish](task-preleveledfinish-property-project.md)|Gets the finish date of a task before leveling occurred. Read-only  **Variant**.|
|[PreleveledStart](task-preleveledstart-property-project.md)|Gets the start date of a task before leveling occurred. Read-only  **Variant**.|
|[Priority](task-priority-property-project.md)|Gets or sets the priority for the task. Read/write  **Variant**.|
|[Project](task-project-property-project.md)|Gets the name of the project containing the  **Task**. Read-only **String**.|
|[RecalcFlags](task-recalcflags-property-project.md)|Gets a bit mask, flagging one or more conditions that are driving the task. Read-only  **Long**.|
|[Recurring](task-recurring-property-project.md)|**True** if the task is a recurring task. Read-only **Variant**.|
|[RegularWork](task-regularwork-property-project.md)|Gets the amount of regular work for the task. Read-only  **Variant**.|
|[RemainingCost](task-remainingcost-property-project.md)|Gets the remaining cost for the task. Read-only  **Variant**.|
|[RemainingDuration](task-remainingduration-property-project.md)|Gets or sets the remaining duration (in minutes) of the task. Read-only for summary tasks. Read/write  **Variant**.|
|[RemainingOvertimeCost](task-remainingovertimecost-property-project.md)|Gets the remaining overtime cost for the task. Read-only  **Variant**.|
|[RemainingOvertimeWork](task-remainingovertimework-property-project.md)|Gets the remaining overtime work (in minutes) for the task. Read-only  **Variant**.|
|[RemainingWork](task-remainingwork-property-project.md)|Gets or sets the remaining work (in minutes) for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[ResourceGroup](task-resourcegroup-property-project.md)|Gets the names of groups associated with the resources assigned to a task, separated by the list separator. Read-only  **String**.|
|[ResourceInitials](task-resourceinitials-property-project.md)|Gets the initials of the resources assigned to a task, separated by the list separator. Read-only  **String**.|
|[ResourceNames](task-resourcenames-property-project.md)|Gets or sets the names of the resources assigned to a task. Read/write  **String**.|
|[ResourcePhonetics](task-resourcephonetics-property-project.md)|Gets the phonetic representation of a resource name. Read-only  **String**.|
|[Resources](task-resources-property-project.md)|Gets a  **[Resources](resource-object-project.md)** collection that contains the resources assigned to the task. Read-only **Resources**.|
|[ResponsePending](task-responsepending-property-project.md)|**True** if a response has not been received for at least one TeamAssign message. Read-only **Boolean**.|
|[Resume](task-resume-property-project.md)|Gets or sets the date that the remaining portion of the task is scheduled to resume after you enter any progress. Read/write  **Variant**.|
|[Rollup](task-rollup-property-project.md)|**True** if the dates of a subtask appear on its corresponding summary task bar. Read/write **Variant**.|
|[ScheduledDuration](task-scheduledduration-property-project.md)|Gets the scheduled (as opposed to actual) duration of a task. Read-only  **Variant**.|
|[ScheduledFinish](task-scheduledfinish-property-project.md)|Gets the scheduled (as opposed to actual) finish time of a task. Read-only  **Variant**|
|[ScheduledStart](task-scheduledstart-property-project.md)|Gets the scheduled (as opposed to actual) start time of a task. Read-only  **Variant**|
|[SPI](task-spi-property-project.md)|Gets the value of the schedule performance index (SPI) calculation for the task. Read-only  **Double**.|
|[SplitParts](task-splitparts-property-project.md)|Gets a  **[SplitParts](splitpart-object-project.md)** collection that represents the portions of a split task. Read-only **SplitParts**.|
|[Start](task-start-property-project.md)|Gets or sets the start date of the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start1](task-start1-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start10](task-start10-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start2](task-start2-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start3](task-start3-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start4](task-start4-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start5](task-start5-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start6](task-start6-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start7](task-start7-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start8](task-start8-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[Start9](task-start9-property-project.md)|Gets or sets a Start local custom field for the task. Read-only for summary tasks. Read/write  **Variant**.|
|[StartDriver](task-startdriver-property-project.md)|Gets the  **[StartDriver](startdriver-object-project.md)** object for the task. Read-only **StartDriver**.|
|[StartSlack](task-startslack-property-project.md)|Gets the starting slack time of a task in minutes. Read-only  **Variant**.|
|[StartText](task-starttext-property-project.md)|Gets or sets a string representation of the task start date. Read/write  **String**.|
|[StartVariance](task-startvariance-property-project.md)|Gets the variance (in minutes) between the baseline start date and the start date of the task. Read-only  **Variant**.|
|[Status](task-status-property-project.md)|Gets the status of a specified task. Read-only  **PjStatusType**.|
|[StatusManagerName](task-statusmanagername-property-project.md)|Gets or sets the GUID of the enterprise resource responsible for accepting or rejecting assignment progress updates for the task. Read/write  **String**.|
|[Stop](task-stop-property-project.md)|Gets or sets the date on which a task stops. Read/write  **Variant**.|
|[Subproject](task-subproject-property-project.md)|Gets or sets the subproject name for the task. Read/write  **String**.|
|[SubProjectReadOnly](task-subprojectreadonly-property-project.md)|**True** if the subproject is read-only. Read/write **Variant**.|
|[Successors](task-successors-property-project.md)|Gets or sets a list of the identification numbers of a task's successors. Read/write  **String**.|
|[SuccessorTasks](task-successortasks-property-project.md)|Gets a  **[Tasks](task-object-project.md)** collection representing the successors of the task. Read-only **Tasks**.|
|[Summary](task-summary-property-project.md)|**True** if the task is a summary task. Read-only **Boolean**.|
|[SV](task-sv-property-project.md)|Gets the earned value scheduled variance (SV) of the task. Read-only  **Variant**.|
|[SVPercent](task-svpercent-property-project.md)|Gets the earned value scheduled variance (SV) percent of the task. Read-only  **Variant**.|
|[TaskDependencies](task-taskdependencies-property-project.md)|Gets a  **[TaskDependencies](taskdependency-object-project.md)** collection of dependent (predecessor and successor) tasks. Read-only **TaskDependencies**.|
|[TCPI](task-tcpi-property-project.md)|Gets the TCPI (to complete performance index) value for the task. Read-only  **Double**.|
|[TeamStatusPending](task-teamstatuspending-property-project.md)|**True** if a response has not been received for at least one progress request message. Read-only **Boolean**.|
|[Text1](task-text1-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text10](task-text10-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text11](task-text11-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text12](task-text12-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text13](task-text13-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text14](task-text14-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text15](task-text15-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text16](task-text16-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text17](task-text17-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text18](task-text18-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text19](task-text19-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text2](task-text2-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text20](task-text20-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text21](task-text21-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text22](task-text22-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text23](task-text23-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text24](task-text24-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text25](task-text25-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text26](task-text26-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text27](task-text27-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text28](task-text28-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text29](task-text29-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text3](task-text3-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text30](task-text30-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text4](task-text4-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text5](task-text5-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text6](task-text6-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text7](task-text7-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text8](task-text8-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[Text9](task-text9-property-project.md)|Gets or sets the value of a local Text custom field for the task. Read/write  **String**.|
|[TotalSlack](task-totalslack-property-project.md)|Gets the total slack time for a task in minutes. Read-only  **Variant**.|
|[Type](task-type-property-project.md)|Gets or sets the way the task is calculated; that is, which one of units, duration, or work are fixed. Read/write  **PjTaskFixedType**.|
|[UniqueID](task-uniqueid-property-project.md)|Gets the unique identification number of the task. Read-only  **Long**.|
|[UniqueIDPredecessors](task-uniqueidpredecessors-property-project.md)|Gets or sets the unique identification ( **UniqueID** ) numbers of the predecessors of a task, separated by the list separator. Read/write **String**.|
|[UniqueIDSuccessors](task-uniqueidsuccessors-property-project.md)|Gets or sets the unique identification ( **UniqueID** ) numbers of the successors of the task, separated by the list separator. Read/write **String**.|
|[UpdateNeeded](task-updateneeded-property-project.md)|**True** if at least one resource assigned to the task needs to be updated regarding the status of the task. Read-only **Boolean**.|
|[VAC](task-vac-property-project.md)|Gets the VAC (Variance At Completion) cost for the task. Read-only  **Variant**.|
|[Warning](task-warning-property-project.md)|Gets the active warning for a task. Read-only  **Variant**.|
|[WBS](task-wbs-property-project.md)|Gets or sets the work breakdown structure (WBS) code of the task. Read/write  **String**.|
|[WBSPredecessors](task-wbspredecessors-property-project.md)|Gets the work breakdown structure (WBS) codes of the task predecessors, separated by the list separator. Read-only  **String**.|
|[WBSSuccessors](task-wbssuccessors-property-project.md)|Gets the work breakdown structure (WBS) codes of the task successors, separated by the list separator. Read-only  **String**.|
|[Work](task-work-property-project.md)|Gets or sets the work (in minutes) for the task. Read/write  **Variant**.|
|[WorkVariance](task-workvariance-property-project.md)|Gets the variance between the baseline work and the work for the task. Read-only  **Variant**.|
|[Compliant](task-compliant-property-project.md)||

