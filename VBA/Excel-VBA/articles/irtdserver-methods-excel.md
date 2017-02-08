---
title: IRtdServer Methods (Excel)
ms.prod: EXCEL
ms.assetid: 80579a8a-bd93-4aa4-bd2a-8d36fc972928
---


# IRtdServer Methods (Excel)

## Methods



|**Name**|**Description**|
|:-----|:-----|
|[ConnectData](irtdserver-connectdata-method-excel.md)|Adds new topics from a real-time data server. The  **ConnectData** method is called when a file is opened that contains real-time data functions or when a user types in a new formula which contains the RTD function.|
|[DisconnectData](irtdserver-disconnectdata-method-excel.md)|Notifies a real-time data (RTD) server application that a topic is no longer in use.|
|[Heartbeat](irtdserver-heartbeat-method-excel.md)|Determines if the real-time data server is still active. Returns a  **Long** value. Zero or a negative number indicates failure; a positive number indicates that the server is active.|
|[RefreshData](irtdserver-refreshdata-method-excel.md)|This method is called by Microsoft Excel to get new data. Returns a  **Variant** .|
|[ServerStart](irtdserver-serverstart-method-excel.md)|The  **ServerStart** method is called immediately after a real-time data server is instantiated. Returns a **Long** ; negative value or zero indicates failure to start the server; positive value indicates success.|
|[ServerTerminate](irtdserver-serverterminate-method-excel.md)|Terminates the connection to the real-time data server.|

