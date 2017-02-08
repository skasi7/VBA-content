---
title: Account Object (Outlook)
keywords: vbaol11.chm3153
f1_keywords:
- vbaol11.chm3153
ms.prod: OUTLOOK
api_name:
- Outlook.Account
ms.assetid: f624438c-4e45-2822-18b6-bfe8074a33c0
---


# Account Object (Outlook)

The  **Account** object represents an account that is defined for the current profile.


## Remarks

The purpose of the [Accounts](accounts-object-outlook.md) collection object and the **Account** object is to provide the capacity to enumerate **Account** objects in a given profile, to identify the type of **Account**, and to use a specific **Account** object to send mail.


 **Note**  Helmut Obertanner provided the following code samples. Helmut is a [Microsoft Most Valuable Professional](https://mvp.microsoft.com/en-us/default.aspx
) with expertise in Microsoft Office development tools in Microsoft Visual Studio and Microsoft Office Outlook.


## Example

The following managed code samples are written in C# and Visual Basic. To run a .NET Framework managed code sample that needs to call into a Component Object Model (COM), you must use an interop assembly that defines and maps managed interfaces to the COM objects in the object model type library. For Outlook, you can use Visual Studio and the Outlook Primary Interop Assembly (PIA). Before you run managed code samples for Outlook 2013, ensure that you have installed the Outlook 2013 PIA and have added a reference to the Microsoft Outlook 15.0 Object Library component in Visual Studio. You should use the following code samples in the  `ThisAddIn` class of an Outlook add-in (using Office Developer Tools for Visual Studio). The **Application** object in the code must be a trusted Outlook **Application** object provided by `ThisAddIn.Globals`. For more information about using the Outlook PIA to develop managed Outlook solutions, see the  **Welcome to the Outlook Primary Interop Assembly Reference** on MSDN.

The following code samples show the  `DisplayAccountInformation` method of the `Sample` class, implemented as part of an Outlook add-in project. Each project adds a reference to the Outlook PIA, which is based on the **Microsoft.Office.Interop.Outlook** namespace. The `DisplayAccountInformation` method takes as an input argument a trusted Outlook[Application](http://msdn.microsoft.com/library/application-object-outlook%28Office.15%29.aspx) object, and uses the **Account** object to display the details of each account that is available for the current Outlook profile.




```C#
using System; 
using System.Text; 
using Outlook = Microsoft.Office.Interop.Outlook; 
 
namespace OutlookAddIn1 
{ 
 class Sample 
 { 
 public static void DisplayAccountInformation(Outlook.Application application) 
 { 
 
 // The Namespace Object (Session) has a collection of accounts. 
 Outlook.Accounts accounts = application.Session.Accounts; 
 
 // Concatenate a message with information about all accounts. 
 StringBuilder builder = new StringBuilder(); 
 
 // Loop over all accounts and print detail account information. 
 // All properties of the Account object are read-only. 
 foreach (Outlook.Account account in accounts) 
 { 
 
 // The DisplayName property represents the friendly name of the account. 
 builder.AppendFormat("DisplayName: {0}\n", account.DisplayName); 
 
 // The UserName property provides an account-based context to determine identity. 
 builder.AppendFormat("UserName: {0}\n", account.UserName); 
 
 // The SmtpAddress property provides the SMTP address for the account. 
 builder.AppendFormat("SmtpAddress: {0}\n", account.SmtpAddress); 
 
 // The AccountType property indicates the type of the account. 
 builder.Append("AccountType: "); 
 switch (account.AccountType) 
 { 
 
 case Outlook.OlAccountType.olExchange: 
 builder.AppendLine("Exchange"); 
 break; 
 
 case Outlook.OlAccountType.olHttp: 
 builder.AppendLine("Http"); 
 break; 
 
 case Outlook.OlAccountType.olImap: 
 builder.AppendLine("Imap"); 
 break; 
 
 case Outlook.OlAccountType.olOtherAccount: 
 builder.AppendLine("Other"); 
 break; 
 
 case Outlook.OlAccountType.olPop3: 
 builder.AppendLine("Pop3"); 
 break; 
 } 
 
 builder.AppendLine(); 
 } 
 
 // Display the account information. 
 System.Windows.Forms.MessageBox.Show(builder.ToString()); 
 } 
 } 
}
```




```VB.net
Imports Outlook = Microsoft.Office.Interop.Outlook 
 
Namespace OutlookAddIn2 
 Class Sample 
 Shared Sub DisplayAccountInformation(ByVal application As Outlook.Application) 
 
 ' The Namespace Object (Session) has a collection of accounts. 
 Dim accounts As Outlook.Accounts = application.Session.Accounts 
 
 ' Concatenate a message with information about all accounts. 
 Dim builder As StringBuilder = New StringBuilder() 
 
 ' Loop over all accounts and print detail account information. 
 ' All properties of the Account object are read-only. 
 Dim account As Outlook.Account 
 For Each account In accounts 
 
 ' The DisplayName property represents the friendly name of the account. 
 builder.AppendFormat("DisplayName: {0}" &amp; vbNewLine, account.DisplayName) 
 
 ' The UserName property provides an account-based context to determine identity. 
 builder.AppendFormat("UserName: {0}" &amp; vbNewLine, account.UserName) 
 
 ' The SmtpAddress property provides the SMTP address for the account. 
 builder.AppendFormat("SmtpAddress: {0}" &amp; vbNewLine, account.SmtpAddress) 
 
 ' The AccountType property indicates the type of the account. 
 builder.Append("AccountType: ") 
 Select Case (account.AccountType) 
 
 Case Outlook.OlAccountType.olExchange 
 builder.AppendLine("Exchange") 
 
 
 Case Outlook.OlAccountType.olHttp 
 builder.AppendLine("Http") 
 
 
 Case Outlook.OlAccountType.olImap 
 builder.AppendLine("Imap") 
 
 
 Case Outlook.OlAccountType.olOtherAccount 
 builder.AppendLine("Other") 
 
 
 Case Outlook.OlAccountType.olPop3 
 builder.AppendLine("Pop3") 
 
 
 End Select 
 
 builder.AppendLine() 
 Next 
 
 
 ' Display the account information. 
 Windows.Forms.MessageBox.Show(builder.ToString()) 
 End Sub 
 
 
 End Class 
End Namespace
```


## Methods



|**Name**|
|:-----|
|[GetAddressEntryFromID](http://msdn.microsoft.com/library/account-getaddressentryfromid-method-outlook%28Office.15%29.aspx)|
|[GetRecipientFromID](http://msdn.microsoft.com/library/account-getrecipientfromid-method-outlook%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AccountType](http://msdn.microsoft.com/library/account-accounttype-property-outlook%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/account-application-property-outlook%28Office.15%29.aspx)|
|[AutoDiscoverConnectionMode](http://msdn.microsoft.com/library/account-autodiscoverconnectionmode-property-outlook%28Office.15%29.aspx)|
|[AutoDiscoverXml](http://msdn.microsoft.com/library/account-autodiscoverxml-property-outlook%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/account-class-property-outlook%28Office.15%29.aspx)|
|[CurrentUser](http://msdn.microsoft.com/library/account-currentuser-property-outlook%28Office.15%29.aspx)|
|[DeliveryStore](http://msdn.microsoft.com/library/account-deliverystore-property-outlook%28Office.15%29.aspx)|
|[DisplayName](http://msdn.microsoft.com/library/account-displayname-property-outlook%28Office.15%29.aspx)|
|[ExchangeConnectionMode](http://msdn.microsoft.com/library/account-exchangeconnectionmode-property-outlook%28Office.15%29.aspx)|
|[ExchangeMailboxServerName](http://msdn.microsoft.com/library/account-exchangemailboxservername-property-outlook%28Office.15%29.aspx)|
|[ExchangeMailboxServerVersion](http://msdn.microsoft.com/library/account-exchangemailboxserverversion-property-outlook%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/account-parent-property-outlook%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/account-session-property-outlook%28Office.15%29.aspx)|
|[SmtpAddress](http://msdn.microsoft.com/library/account-smtpaddress-property-outlook%28Office.15%29.aspx)|
|[UserName](http://msdn.microsoft.com/library/account-username-property-outlook%28Office.15%29.aspx)|

## See also


#### Other resources


[Account Object Members](http://msdn.microsoft.com/library/account-members-outlook%28Office.15%29.aspx)
[How to: Send an E-mail Given the SMTP Address of an Account](http://msdn.microsoft.com/library/send-an-e-mail-given-the-smtp-address-of-an-account-outlook%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/object-model-outlook-vba-reference%28Office.15%29.aspx)
