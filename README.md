<div align="center">

## How to destroy a Shopping Cart


</div>

### Description

I often get questions on how to deal with shopping cart data stored within the database that is no longer needed after the user has logged off. The answer is simple and deals with some database calls in the Global.ASA.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-how-to-destroy-a-shopping-cart__4-6632/archive/master.zip)





### Source Code

<P>
As you user adds an item into there shopping cart, you should create a field that also holds the value of the SessionID (Session.SessionID). This ID is unique to all of your "current" visitors. When a users session ends, you can "clean up" your database by deleting all of the records relating to that session ID. Here is a brief run-down on the SQL you would execute in your Global.ASA file.
</P>
<DIV><FONT size=2><FONT face=Courier><FONT color=#0000ff><</FONT><FONT
color=#800000>SCRIPT</FONT> <FONT color=#ff0000>language</FONT><FONT
color=#0000ff>="vbScript"</FONT> <FONT color=#ff0000>runat</FONT><FONT
color=#0000ff>="Server"></FONT></FONT></FONT></DIV>
<DIV><FONT face=Courier size=2></FONT> </DIV><FONT size=2>
<DIV><FONT size=2><FONT face=Courier><STRONG><FONT color=#0000ff>Sub
</FONT></STRONG>Session_onStart()</FONT></FONT></DIV>
<DIV><FONT face=Courier size=2>    lStrSQL = <FONT color=#ff0000>"DELETE FROM
ShoppingCartItems WHERE SessionID = "</FONT> &
Session.SessionID</FONT></DIV>
<DIV><FONT face=Courier color=#0000ff size=2><STRONG>End
Sub</STRONG></FONT></DIV>
<DIV><FONT face=Courier></FONT> </DIV>
<DIV><FONT face=Courier color=#008000>' this is the main one you want - the
others just help make </FONT></DIV>
<DIV><FONT face=Courier color=#008000>' sure records get deleted if they are not
needed.</FONT></DIV>
<DIV><FONT face=Courier><FONT color=#000080><STRONG><FONT
color=#0000ff>Sub</FONT> </STRONG></FONT>Session_onEnd()</FONT></FONT></DIV>
<DIV><FONT face=Courier size=2>    lStrSQL = <FONT color=#ff0000>"DELETE FROM
ShoppingCartItems WHERE SessionID = "</FONT> &
Session.SessionID</FONT></DIV>
<DIV><FONT face=Courier color=#0000ff size=2><STRONG>End
Sub</STRONG></FONT></DIV>
<DIV><FONT face=Courier size=2></FONT> </DIV>
<DIV><FONT size=2><FONT face=Courier><STRONG><FONT color=#0000ff>Sub
</FONT></STRONG>Application_onStart()</FONT></FONT></DIV>
<DIV><FONT size=2><FONT face=Courier>    lStrSQL = <FONT color=#ff0000>"DELETE
FROM ShoppingCartItems"</FONT></FONT></FONT></DIV>
<DIV><FONT face=Courier color=#0000ff size=2><STRONG>End
Sub</STRONG></FONT></DIV>
<DIV><FONT face=Courier size=2></FONT> </DIV>
<DIV><FONT size=2>
<DIV><FONT size=2><FONT face=Courier><STRONG><FONT
color=#0000ff>Sub</FONT></STRONG> Application_onEnd()</FONT></FONT></DIV>
<DIV><FONT size=2><FONT face=Courier>    lStrSQL = <FONT color=#ff0000>"DELETE
FROM ShoppingCartItems"</FONT></FONT></FONT></DIV>
<DIV><FONT face=Courier color=#0000ff size=2><STRONG>End
Sub</STRONG></FONT></DIV></FONT></DIV>
<DIV><FONT color=#800000 size=2><FONT face=Courier><FONT
color=#0000ff></</FONT>SCRIPT<FONT
color=#0000ff>></FONT></FONT></FONT></DIV>

