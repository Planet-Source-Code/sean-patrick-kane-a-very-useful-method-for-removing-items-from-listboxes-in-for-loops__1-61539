<div align="center">

## A very useful method for removing items from listboxes in for loops


</div>

### Description

Many people, at some point in their VB programming careers, will use a for-loop to remove items from a listbox. I would say that easily 90% of the people that code for a situation like this are coding it INCORRECTLY and are using extremely poor programming style. I would encourage beginners and experts alike to examine this article to learn the best way to use a for loop to remove listbox items.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sean Patrick Kane](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-patrick-kane.md)
**Level**          |Beginner
**User Rating**    |4.5 (27 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sean-patrick-kane-a-very-useful-method-for-removing-items-from-listboxes-in-for-loops__1-61539/archive/master.zip)





### Source Code

<font size="6"><b>Before We Begin</b></font>
<br>This is something that every programmer should learn! The situation I'm going to present assumes that we have one listbox on our form, named List1, that has any number of items within it. For whatever reason, we want to loop through the listbox, and remove certain items -- say, any items that have the letter "B" in it; however, please realize, a much more common use of this code would be to search for duplicate items, or items that match a specific need or requirement. I'm going to show you the common way of doing things, which is extremely ugly and incorrect, and then I'm going to demonstrate the proper way to go
about the situation I have presented.
<p><font size="6"><b>The Wrong (but most common) Method</b></font>
<br>Before any programmer attempts to debug their application, they may write something like this:
<p><b>Dim i as integer 'Declare our variables
<br>For i = 0 to List1.ListCount - 1 'Loop through the listbox
<br><ul>If InStr(List1.List(i), "B") <> 0 Then  'B exists in this item...we need to remove it
<br>List1.RemoveItem i 'Remove the offending item
<br>End If
<br></ul>Next i</b>
<p>Unfortunately, once the program ran, the coder would be presented with an error message saying "Run-time error '13': Type mismatch". What does this programmer do? All too often, they just throw in a line similar to "On Error Resume Next", and magically the error message goes away, and usually it doesn't affect the actual function of the problematic function. This is, of course, a really bad way to troubleshoot a problem. Before fixing the problem, we should understand what the problem actually is...
<p><font size="6"><b>Why The Error Message?</b></font>
<br>If the coder would use the debug stuff in VB, they would find that their program is trying to address an index (the i variable) that no longer exists -- often times the i variable equals the List1.ListCount. Basically what has happened is that the For loop stores the ListCount when it starts, and never re-initializes it. For instance, if List1 contains 5 items before the for-loop runs, and during the loop we remove 2 items, the for-loop would still be running as "For i = 0 to 5", rather than "For i = 0 to 3" (since we removed two items). This of course will cause an error once we try to access List1.List(4) or List1.List(5) because those items will not exist. Now that we know the problem, we can rid ourselves of that nasty "On Error Resume Next" line and actually code the thing correctly.
<p><font size="6"><b>The Correct (but unfortunately uncommon) Method</b></font>
<br>There is a very, very simple way to correct this problem. Instead of having our for-loop start at the beginning of the listbox and move to the end, all we have to do is start at the end and move to the beginning. If we code the for-loop this way, we will never try to access an invalid index. The corrected code is shown below:
<p><b>Dim i as integer 'Declare our variables
<br>For i = (List1.ListCount - 1) to 0 Step -1 'Loop through the listbox backwards
<br><ul>If InStr(List1.List(i), "B") <> 0 Then  'B exists in this item...we need to remove it
<br>List1.RemoveItem i 'Remove the offending item
<br>End If
<br></ul>Next i</b>
<p><font size="6"><b>Conclusion</b></font>
<br>Using the new method, any number of listbox items can be removed safely without the problem of addressing invalid index items. I realize this is a fairly basic concept, but I've seen too many PSC submissions that were using the incorrect method. Any comments or votes would be greatly appreciated.

