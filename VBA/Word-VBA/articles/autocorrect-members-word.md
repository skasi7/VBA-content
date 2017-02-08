---
title: AutoCorrect Members (Word)
ms.prod: WORD
ms.assetid: cc5f42d4-6689-221f-5ad2-3b56f3b2c42f
---


# AutoCorrect Members (Word)
Represents the AutoCorrect functionality in Word.

Represents the AutoCorrect functionality in Word.


## Properties



|**Name**|**Description**|
|:-----|:-----|
|[Application](autocorrect-application-property-word.md)|Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.|
|[CorrectCapsLock](autocorrect-correctcapslock-property-word.md)| **True** if Word automatically corrects instances in which you use the CAPS LOCK key inadvertently as you type. Read/write **Boolean** .|
|[CorrectDays](autocorrect-correctdays-property-word.md)| **True** if Word automatically capitalizes the first letter of days of the week. Read/write **Boolean** .|
|[CorrectHangulAndAlphabet](autocorrect-correcthangulandalphabet-property-word.md)| **True** if Microsoft Word automatically applies the correct font to Latin words typed in the middle of Hangul text or vice versa. Read/write **Boolean** .|
|[CorrectInitialCaps](autocorrect-correctinitialcaps-property-word.md)| **True** if Word automatically makes the second letter lowercase if the first two letters of a word are typed in uppercase. For example, "WOrd" is corrected to "Word." Read/write **Boolean** .|
|[CorrectKeyboardSetting](autocorrect-correctkeyboardsetting-property-word.md)| **True** if Microsoft Word automatically transposes words to their native alphabet if you type text in a language other than the current keyboard language. Read/write **Boolean** .|
|[CorrectSentenceCaps](autocorrect-correctsentencecaps-property-word.md)| **True** if Word automatically capitalizes the first letter in each sentence. Read/write **Boolean** .|
|[CorrectTableCells](autocorrect-correcttablecells-property-word.md)| **True** to automatically capitalize the first letter of table cells. Read/write **Boolean** .|
|[Creator](autocorrect-creator-property-word.md)|Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long** .|
|[DisplayAutoCorrectOptions](autocorrect-displayautocorrectoptions-property-word.md)| **True** for Microsoft Word to display the **AutoCorrect Options** button. Read/write **Boolean** .|
|[Entries](autocorrect-entries-property-word.md)|Returns an  **[AutoCorrectEntries](autocorrectentries-object-word.md)** collection that represents the current list of AutoCorrect entries.|
|[FirstLetterAutoAdd](autocorrect-firstletterautoadd-property-word.md)| **True** if Word automatically adds abbreviations to the list of AutoCorrect First Letter exceptions. Read/write **Boolean** .|
|[FirstLetterExceptions](autocorrect-firstletterexceptions-property-word.md)|Returns a  **[FirstLetterExceptions](firstletterexceptions-object-word.md)** collection that represents the list of abbreviations after which Word won't automatically capitalize the next letter. Read-only.|
|[HangulAndAlphabetAutoAdd](autocorrect-hangulandalphabetautoadd-property-word.md)| **True** if Microsoft Word automatically adds words to the list of Hangul and alphabet AutoCorrect exceptions. Read/write **Boolean** .|
|[HangulAndAlphabetExceptions](autocorrect-hangulandalphabetexceptions-property-word.md)|Returns a  **[HangulAndAlphabetExceptions](hangulandalphabetexceptions-object-word.md)** collection that represents the list of Hangul and alphabet AutoCorrect exceptions.|
|[OtherCorrectionsAutoAdd](autocorrect-othercorrectionsautoadd-property-word.md)| **True** if Microsoft Word automatically adds words to the list of AutoCorrect exceptions on the **Other Corrections** tab in the **AutoCorrect Exceptions** dialog box ( **AutoCorrect Options** command, **Tools** menu). Word adds a word to this list if you delete and then retype a word that you didn't want Word to correct. Read/write **Boolean** .|
|[OtherCorrectionsExceptions](autocorrect-othercorrectionsexceptions-property-word.md)|Returns an  **[OtherCorrectionsExceptions](othercorrectionsexceptions-object-word.md)** collection that represents the list of words that Microsoft Word won't correct automatically.|
|[Parent](autocorrect-parent-property-word.md)|Returns an  **Object** that represents the parent object of the specified **AutoCorrect** object.|
|[ReplaceText](autocorrect-replacetext-property-word.md)| **True** if Microsoft Word automatically replaces specified text with entries from the AutoCorrect list. Read/write **Boolean** .|
|[ReplaceTextFromSpellingChecker](autocorrect-replacetextfromspellingchecker-property-word.md)| **True** if Microsoft Word automatically replaces misspelled text with suggestions from the spelling checker as the user types. Word only replaces words that contain a single misspelling and for which the spelling dictionary only lists one alternative. Read/write **Boolean** .|
|[TwoInitialCapsAutoAdd](autocorrect-twoinitialcapsautoadd-property-word.md)| **True** if Microsoft Word automatically adds words to the list of AutoCorrect Initial Caps exceptions. A word is added to this list if you delete and then retype the uppercase letter (following the initial uppercase letter) that Word changed to lowercase. Read/write **Boolean** .|
|[TwoInitialCapsExceptions](autocorrect-twoinitialcapsexceptions-property-word.md)|Returns a  **[TwoInitialCapsExceptions](twoinitialcapsexceptions-object-word.md)** collection that represents the list of terms containing mixed capitalization that Word won't correct automatically.|

