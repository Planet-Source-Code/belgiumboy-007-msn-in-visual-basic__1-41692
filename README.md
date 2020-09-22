<div align="center">

## \_MSN In Visual Basic\_

<img src="PIC200212181241224772.JPG">
</div>

### Description

This article will cover the most important aspects of programming MSN using Visual Basic.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-12-17 20:16:04
**By**             |[BelgiumBoy\_007](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/belgiumboy-007.md)
**Level**          |Beginner
**User Rating**    |4.9 (151 globes from 31 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[\_MSN\_In\_Vi15153412182002\.zip](https://github.com/Planet-Source-Code/belgiumboy-007-msn-in-visual-basic__1-41692/archive/master.zip)





### Source Code

<html>
<head>
<style type="text/css">
body,table {font:9pt verdana}
h1 {font:20pt verdana}
h2 {font:15pt verdana}
pre {background:EEEEEE}
A:hover {COLOR: gold; TEXT-DECORATION: none}
A:hover {font-weight: bold}
</style>
<title></title>
</head>
<body link="#0000FF" vlink="#0000FF" alink="#0000FF">
<h1>Programming MSN Messenger using Visual Basic.</h1>
<table width="700">
 <tr>
  <td>This article will cover the more simple actions that can be done.&nbsp; Here is a list
  of the different topics:<ul>
   <li><a href="#intro">Intro.</a></li>
   <li><a href="#whatdoesvisualbasicallowmetodo">What does Visual Basic allow me to do ?</a></li>
   <li><a href="#declaringtheappropriatevariables">Declaring the appropriate variables.</a></li>
   <li><a href="#signinginandoutoftheservice">Signing in to/out of the .NET Messenger Service.</a></li>
   <li><a href="#changeyournickname">Change your nickname.</a></li>
   <li><a href="#changeyourstatus">Change your status.</a></li>
   <li><a href="#sendyourcontactslisttoalistbox">Send your contacts list to a listbox.</a></li>
   <li><a href="#sendyourcontactslisttoatreeview">Send your contacts list to a treeview.</a></li>
   <li><a href="#sendaninstantmessage">Send an Instant Message.</a></li>
   <li><a href="#openanimwindow">Open an IM window.</a></li>
   <li><a href="#blockunblockcontacts">Block/unblock contacts.</a></li>
   <li><a href="#addremovecontacts">Add/remove contacts.</a></li>
   <li><a href="#handlingmsnevents">Handling MSN events.</a></li>
   <li><a href="#creatingaselfupdatingcontactslist">Creating a self-updating contacts list.</a></li>
  </ul>
  <p>If you want the code for a fully functioning MSN-Bot then click <a
  href="http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=40655&amp;lngWId=1"
  target="_blank">here</a>.</td>
 </tr>
</table>
<h1><a name="intro"></a>Intro.</h1>
<table border="0" width="700">
 <tr>
  <td>So you want to know how to program MSN Messenger using Microsoft Visual Basic? &nbsp;
  Well you've come to the right place.&nbsp; I'll take you thru the basics.&nbsp; I'm going
  to assume that you already know some Visual Basic so I'll concentrate on explaining the
  MSN-related code.&nbsp; I hope this article will help you and if you have any questions
  than please don't hesitate to <a
  href="mailto:webmaster@bartnet.freeservers.com?subject=Question About MSN In Visual Basic">ask
  me</a>.<p><font color="#FF0000">IMPORTANT</font> : The codes here will not work with the
  new version of MSN Messenger (5.0), only with versions 4.7 or below.&nbsp; The ideal
  version is 4.6.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="whatdoesvisualbasicallowmetodo"></a>What does Visual Basic allow me to do ?</h1>
<table border="0" width="700">
 <tr>
  <td>With Visual Basic you can basically do anything that MSN Messenger can do.&nbsp; Al
  you do using Visual Basic is send commands to the MSN program, which is why your program
  will not work if MSN Messenger version 4.7 or below is not running.&nbsp; You do not have
  to be signed in to the .NET Messenger Service but the program must be running.&nbsp; This
  means that you can send/receive Instant Messages, add/remove contacts, change your
  nickname, change your status, go to you e-mail inbox, ...<p><font color="#000000"><a
  href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="declaringtheappropriatevariables"></a>Declaring the appropriate variables.</h1>
<table border="0" width="700">
 <tr>
  <td>First you have to add the Messenger references to your project.&nbsp; Go to <u>P</u>rojects
  &gt; Prefere<u>n</u>ces... and select the following references:<ul>
   <li>Messenger Type Library
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    (found at C:\Program Files\Messenger\msmsgs.exe)</li>
   <li>Messenger API Type Library
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (found at
    C:\Program Files\Messenger\msmsgs.exe\3)</li>
   <li>Messenger AddIns Type Library&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (found at
    C:\Program Files\Messenger\msmsgs.exe\4)</li>
   <li>Messenger Private Type Library&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (found at
    C:\Program Files\Messenger\msmsgs.exe\2)</li>
  </ul>
  <p>You most likely won't need all of them but it's best to add them anyway.&nbsp; Now you
  need to declare the standard variable.&nbsp; There are 2 that we will be using in this
  tutorial, the second one can be declared in two different ways.&nbsp; The first one is a
  variable using the API Type Library.&nbsp; We'll call it MSNAPI:</p>
  <pre><font color="#000080">Private</font> MSNAPI <font color="#000080">As New</font> MessengerAPI.Messenger</pre>
  <p>The second one will use the Messenger Type Library, we'll call this one MSN:</p>
  <pre><font color="#000080">Private</font> MSN <font color="#000080">As New</font> MsgrObject</pre>
  <p>with both of these declarations you can leave out the <font color="#000080">New</font>
  but then you'll have to add this to your code:</p>
  <pre><font color="#000080">Private Sub</font> Form_Load()
 <font color="#000080">  Set</font> MSN = <font
color="#000080">New</font> MsgrObject
<font color="#000080">  Set</font> MSNAPI = <font
color="#000080">New</font> MessengerAPI.Messenger
<font color="#000080">End Sub</font></pre>
  <p>Why would we use the second method ?&nbsp; Simple, when we want to add the WithEvents
  option we have to leave it out.&nbsp; Most of the time you are going to be using the
  second method because then you can keep track of what's happening in MSN and even block
  most of the events.&nbsp; Then the declaration would look like this:</p>
  <pre><font color="#000080">Private WithEvents</font> MSN <font color="#000080">As</font> MsgrObject
<font
color="#000080">Private WithEvents</font> MSNAPI <font color="#000080">As</font> MessengerAPI.Messenger</pre>
  <p>But for this article we will use the first method simply because it's shorter.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="signinginandoutoftheservice"></a>Signing in to/out of the .NET Messenger
Service.</h1>
<table border="0" width="700">
 <tr>
  <td>This is very simple.&nbsp; Create a new project, add the references and on the form
  put 2 Labels (lblUserEmail &amp; lblPassword), 2 TextBoxes (txtUserEmail &amp;
  txtPassword) and 2 CommandButtons (cmdSignIn &amp; cmdSignOut).&nbsp; Now enter the
  following code:<pre><font color="#000080">Private</font> MSN <font color="#000080">As New</font> MsgrObject
<font
color="#000080">Private Sub</font> cmdSignIn_Click()
<font color="#000080">On Error Resume Next</font>
  MSN.Logon txtUserEmail.Text, txtPassword.Text, MSN.Services.PrimaryService
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdSignOut_Click()
<font
color="#000080">On Error Resume Next</font>
  MSN.Logoff
<font color="#000080">End Sub</font></pre>
  <p>As you can see it isn't hard at all.&nbsp; I've put in error handling because if you
  try to Sign In when you are already signed in you will get an error and the same with
  signing out.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 1</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="changeyournickname"></a>Change your nickname.</h1>
<table border="0" width="700">
 <tr>
  <td>It only takes one line of code but unfortunately because you are using the running MSN
  Messenger program you can't change your nickname to everything you want.&nbsp; Rude words,
  websites and some other words will not work.&nbsp; Create a new project, add the
  references and on the form put a Label (lblNewNickName), a TextBox (txtNewNickName) and a
  CommandButton (cmdChangeNickName).&nbsp; Enter the following code:<pre><font
color="#000080">Private</font> MSN <font color="#000080">As</font> <font color="#000080">New</font> MsgrObject
<font
color="#000080">Private</font> <font color="#000080">Sub</font> cmdChangeNickName_Click()
<font
color="#000080">  If</font> MSN.LocalState = MSTATE_OFFLINE <font color="#000080">Then</font>
    MsgBox &quot;You are not Signed In&quot;
<font
color="#000080">  Else</font>
    MSN.Services.PrimaryService.FriendlyName = txtNewNickName.Text
    txtNewNickName.Text = &quot;&quot;
<font
color="#000080">  End</font> <font color="#000080">If</font>
<font color="#000080">End</font> <font
color="#000080">Sub</font></pre>
  <p>You may have noticed that the following also exists:</p>
  <pre>MSN.LocalFriendlyName</pre>
  <p>This is read only so it will not work.&nbsp; When a user presses the button then first
  the program will make sure that you are Signed In and if you aren't it will give a Message
  Box.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 2</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="changeyourstatus"></a>Change your status.</h1>
<table border="0" width="700">
 <tr>
  <td>This is also very easy.&nbsp; Create a new project, add the references and on the form
  put 7 OptionButtons (optOnline, optBusy, optBeRightBack, optAway, optOnThePhone,
  optOutToLunch &amp; optAppearOffline).&nbsp; Next, insert the following code:<pre><font
color="#000080">Private</font> MSN <font color="#000080">As New</font> MsgrObject
<font
color="#000080">Private Sub</font> Form_Load()
  <font color="#000080">Select Case</font> MSN.LocalState
    <font
color="#000080">Case</font> MSTATE_ONLINE
      optOnline.Value = <font color="#000080">True</font>
    <font
color="#000080">Case</font> MSTATE_BUSY
      optBusy.Value = <font color="#000080">True</font>
    <font
color="#000080">Case</font> MSTATE_BE_RIGHT_BACK
      optBeRightBack.Value = <font
color="#000080">True</font>
    <font color="#000080">Case</font> MSTATE_AWAY
      optAway.Value = <font
color="#000080">True</font>
    <font color="#000080">Case</font> MSTATE_ON_THE_PHONE
      optOnThePhone.Value = <font
color="#000080">True</font>
    <font color="#000080">Case</font> MSTATE_OUT_TO_LUNCH
      optOutToLunch.Value = <font
color="#000080">True</font>
    <font color="#000080">Case</font> MSTATE_INVISIBLE
      optAppearOffline.Value = <font
color="#000080">True</font>
  <font color="#000080">End Select</font>
<font color="#000080">End Sub</font>
<font
color="#000080">Private Sub</font> optAppearOffline_Click()
  MSN.LocalState = MSTATE_INVISIBLE
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> optAway_Click()
  MSN.LocalState = MSTATE_AWAY
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> optBeRightBack_Click()
  MSN.LocalState = MSTATE_BE_RIGHT_BACK
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> optBusy_Click()
  MSN.LocalState = MSTATE_BUSY
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> optOnline_Click()
  MSN.LocalState = MSTATE_ONLINE
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> optOnThePhone_Click()
  MSN.LocalState = MSTATE_ON_THE_PHONE
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> optOutToLunch_Click()
  MSN.LocalState = MSTATE_OUT_TO_LUNCH
<font
color="#000080">End Sub</font></pre>
  <p>When you run the program it will first check what your status is and set the
  appropriate OptionButton value to true.&nbsp; Then depending on which OptionButton has
  been clicked it will change your status.&nbsp; Here's a list of all of the MSN.LocalState
  constants:<ul>
   <li>MSTATE_AWAY</li>
   <li>MSTATE_BE_RIGHT_BACK</li>
   <li>MSTATE_BUSY</li>
   <li>MSTATE_IDLE</li>
   <li>MSTATE_INVISIBLE</li>
   <li>MSTATE_LOCAL_CONNECTING_TO_SERVER</li>
   <li>MSTATE_LOCAL_DISCONNECTING_FROM_SERVER</li>
   <li>MSTATE_LOCAL_FINDING_SERVER</li>
   <li>MSTATE_LOCAL_SYNCHRONIZING_WITH_SERVER</li>
   <li>MSTATE_OFFLINE</li>
   <li>MSTATE_ON_THE_PHONE</li>
   <li>MSTATE_ONLINE</li>
   <li>MSTATE_OUT_TO_LUNCH</li>
   <li>MSTATE_UNKNOWN</li>
  </ul>
  <p><font color="#FF0000">The code for this example can be found in folder number 3</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="sendyourcontactslisttoalistbox"></a>Send your contacts list to a listbox.</h1>
<table border="0" width="700">
 <tr>
  <td>This is a bit more complicated than the last 3 topics.&nbsp; Create a new project, add
  the references and on the form put 2 Labels (lblOnlineContacts &amp; lblOfflineContacts),
  2 ListBoxes (lstOnlineContacts &amp; lstOfflineContacts) and a CommandButton
  (cmdRefreshList).&nbsp; Now insert the following code:<pre><font color="#000080">Private</font> MSN <font
color="#000080">As New</font> MsgrObject
<font color="#000080">Private Sub</font> RefreshList()
  lstOfflineContacts.Visible = <font
color="#000080">False</font>
  lstOnlineContacts.Visible = <font color="#000080">False</font>
  <font
color="#000080">Dim</font> User <font color="#000080">As</font> IMsgrUser
  lstOnlineContacts.Clear
  lstOfflineContacts.Clear
  <font
color="#000080">For Each</font> User <font color="#000080">In</font> MSN.List(MLIST_CONTACT)
    <font
color="#000080">If</font> User.State = MSTATE_OFFLINE <font color="#000080">Then</font>
      lstOfflineContacts.AddItem (User.EmailAddress)
    <font
color="#000080">Else</font>
      lstOnlineContacts.AddItem (User.EmailAddress)
    <font
color="#000080">End If</font>
  <font color="#000080">Next</font>
  lstOfflineContacts.Visible = <font
color="#000080">True</font>
  lstOnlineContacts.Visible = <font color="#000080">True</font>
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdRefreshList_Click()
  <font
color="#000080">If</font> MSN.LocalState &lt;&gt; MSTATE_OFFLINE <font color="#000080">Then</font> RefreshList
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> Form_Load()
  cmdRefreshList_Click
<font
color="#000080">End Sub</font></pre>
  <p>The most important part of this code is the RefreshList() sub.&nbsp; Here we first hide
  both the ListBoxes and at the end show them again so there's no flickering.&nbsp; Next we
  declare another variable (User), you can set this variable to any user and get their
  nickname, send a message, ... .&nbsp; In the For - Next loop we go thru all the contacts
  and check their current status.&nbsp; Of they are offline then we add them to the offline
  contacts ListBox, otherwise the offline contacts ListBox.&nbsp; In this code the e-mail
  address of every contact is added to the list but you can also add the nickname of every
  contact, just change User.EmailAddress to User.FriendlyName.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 4</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="sendyourcontactslisttoatreeview"></a>Send your contacts list to a treeview.</h1>
<table border="0" width="700">
 <tr>
  <td>This is the more advanced way of displaying your contacts list, this will look almost
  identical to the way that MSN displays your list.&nbsp; Create a new project, add the
  references, add Microsoft Windows Common Controls 6.0 (SP4) (found in
  C:\Windows\System32\mscomctl.ocx) and on the form put a Label (lblContacts), a TreeView
  (tContacts), a CommandButton (cmdRefreshList) and an ImageList (ilsIcons).&nbsp; Add all
  the picture files found in folder number 5 to the ImageList and use the exact filename
  without the extension as the key (keep in mind that it's case sensitive).&nbsp; Now insert
  the following code into your form:<pre><font color="#000080">Private</font> MSN <font
color="#000080">As New</font> MsgrObject
<font color="#000080">Private Sub</font> RefreshList()
  tContacts.Visible = <font
color="#000080">False</font>
  tContacts.Nodes.Clear
  <font color="#000080">Dim</font> User <font
color="#000080">As</font> IMsgrUser
  <font color="#000080">Dim</font> UsersOnline <font
color="#000080">As</font> <font color="#000080">Integer</font>
  <font color="#000080">Dim</font> UsersOffline <font
color="#000080">As</font> <font color="#000080">Integer</font>
  <font color="#000080">For Each</font> User <font
color="#000080">In</font> MSN.List(MLIST_CONTACT)
    <font color="#000080">If</font> User.State = MSTATE_OFFLINE <font
color="#000080">Then</font>
      UsersOffline = UsersOffline + 1
    <font
color="#000080">Else</font>
      UsersOnline = UsersOnline + 1
    <font
color="#000080">End If</font>
  <font color="#000080">Next</font>
  tContacts.Nodes.Add , , &quot;Online&quot;, &quot;Online (&quot; &amp; UsersOnline &amp; &quot;)&quot;, &quot;Up&quot;, &quot;UpSelected&quot;
  <font
color="#000080">With</font> tContacts.Nodes(1)
    .Selected = <font color="#000080">False</font>
    .Expanded = <font
color="#000080">True</font>
    .Bold = <font color="#000080">True</font>
    .ForeColor = &amp;H8000000D
    .Sorted = <font
color="#000080">True</font>
  <font color="#000080">End With</font>
  tContacts.Nodes.Add , , &quot;Offline&quot;, &quot;Offline (&quot; &amp; UsersOffline &amp; &quot;)&quot;, &quot;Down&quot;, &quot;DownSelected&quot;
  <font
color="#000080">With</font> tContacts.Nodes(2)
    .Selected = <font color="#000080">True</font>
    .Expanded = <font
color="#000080">False</font>
    .Bold = <font color="#000080">True</font>
    .ForeColor = &amp;H8000000D
    .Sorted = <font
color="#000080">True</font>
  <font color="#000080">End With</font>
  <font
color="#000080">For Each</font> User <font color="#000080">In</font> MSN.List(MLIST_ALLOW)
    <font
color="#000080">Select Case</font> User.State
      <font color="#000080">Case</font> MSTATE_AWAY
        tContacts.Nodes.Add &quot;Online&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot; (Away)&quot;, &quot;Away&quot;, &quot;AwaySelected&quot;
      <font
color="#000080">Case</font> MSTATE_BE_RIGHT_BACK
        tContacts.Nodes.Add &quot;Online&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot; (Be Right Back)&quot;, &quot;Away&quot;, &quot;AwaySelected&quot;
      <font
color="#000080">Case</font> MSTATE_BUSY
        tContacts.Nodes.Add &quot;Online&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot; (Busy)&quot;, &quot;Busy&quot;, &quot;BusySelected&quot;
      <font
color="#000080">Case</font> MSTATE_OFFLINE
        tContacts.Nodes.Add &quot;Offline&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot;Offline&quot;, &quot;OfflineSelected&quot;
      <font
color="#000080">Case</font> MSTATE_ON_THE_PHONE
        tContacts.Nodes.Add &quot;Online&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot; (On The Phone)&quot;, &quot;Busy&quot;, &quot;BusySelected&quot;
      <font
color="#000080">Case</font> MSTATE_ONLINE
        tContacts.Nodes.Add &quot;Online&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot;Online&quot;, &quot;OnlineSelected&quot;
      <font
color="#000080">Case</font> MSTATE_OUT_TO_LUNCH
        tContacts.Nodes.Add &quot;Online&quot;, tvwChild, User.EmailAddress, User.FriendlyName, &amp; _
        &quot; (Out To Lunch)&quot;, &quot;Away&quot;, &quot;AwaySelected&quot;
    <font
color="#000080">End Select</font>
  <font color="#000080">Next</font>
  <font
color="#000080">For Each</font> User <font color="#000080">In</font> MSN.List(MLIST_BLOCK)
    <font
color="#000080">If</font> User.State = MSTATE_OFFLINE <font color="#000080">Then</font>
      tContacts.Nodes.Add &quot;Offline&quot;, tvwChild, User.EmailAddress, User.FriendlyName
      &amp; &quot; (Blocked)&quot;, &quot;Blocked&quot;, &quot;BlockedSelected&quot;
    <font
color="#000080">Else</font>
      tContacts.Nodes.Add &quot;Offline&quot;, tvwChild, User.EmailAddress, User.FriendlyName
      &amp; &quot; (Blocked)&quot;, &quot;Blocked&quot;, &quot;BlockedSelected&quot;
    <font
color="#000080">End If</font>
  <font color="#000080">Next</font>
  tContacts.Visible = True
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdRefreshList_Click()
  RefreshList
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> Form_Load()
  RefreshList
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> tContacts_NodeClick(<font
color="#000080">ByVal</font> Node <font color="#000080">As</font> MSComctlLib.Node)
  <font
color="#000080">If</font> Node.Key = &quot;Offline&quot; <font color="#000080">Or</font> Node.Key = &quot;Online&quot; <font
color="#000080">Then</font>
    <font color="#000080">If</font> Node.Expanded = <font
color="#000080">True</font> <font color="#000080">Then</font>
      Node.Expanded = <font
color="#000080">False</font>
      Node.Image = &quot;Down&quot;
      Node.SelectedImage = &quot;DownSelected&quot;
    <font
color="#000080">Else</font>
      Node.Expanded = <font color="#000080">True</font>
      Node.Image = &quot;Up&quot;
      Node.SelectedImage = &quot;UpSelected&quot;
    <font
color="#000080">End If</font>
  <font color="#000080">Else</font>
  <font
color="#000080">End If</font>
<font color="#000080">End Sub</font></pre>
  <p>Here the first thing the program will do is see how many of your contacts are online/offline
  and it will add the 2 tags, then it formats the tags to make them look good. &nbsp; After
  that it goes thru all of the contacts and adds them to the list, they will have a
  different caption and image depending on their status.&nbsp; Notice how I put in some code
  to handle the NodeClick event.&nbsp; This is so that the arrow changes from up to down or
  down to up.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 5</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="sendaninstantmessage"></a>Send an Instant Message.</h1>
<table border="0" width="700">
 <tr>
  <td>To do this we will use some of the previous code.&nbsp; Create a new project, add the
  references and on the form put a Label (lblOnlineContacts), a ListBox (lstOnlineContacts)
  and 2 CommandButtons (cmdRefreshList &amp; cmdSendIM).&nbsp; Insert the following code:<pre><font
color="#000080">Private</font> MSN <font color="#000080">As New</font> MsgrObject
<font
color="#000080">Private Sub</font> RefreshList()
  lstOnlineContacts.Visible = <font
color="#000080">False</font>
  <font color="#000080">Dim</font> User <font color="#000080">As</font> IMsgrUser
  lstOnlineContacts.Clear
  <font
color="#000080">For Each</font> User <font color="#000080">In</font> MSN.List(MLIST_CONTACT)
    <font
color="#000080">If</font> User.State &lt;&gt; MSTATE_OFFLINE <font color="#000080">Then</font> lstOnlineContacts.AddItem (User.EmailAddress)
  <font
color="#000080">Next</font>
  lstOnlineContacts.Visible = <font color="#000080">True</font>
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdRefreshList_Click()
  <font
color="#000080">If</font> MSN.LocalState &lt;&gt; MSTATE_OFFLINE <font color="#000080">Then</font> RefreshList
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdSendIM_Click()
  <font
color="#000080">Dim</font> User <font color="#000080">As</font> IMsgrUser
  <font
color="#000080">Dim</font> bstrMsgHeader <font color="#000080">As</font> String
  <font
color="#000080">Dim</font> bstrMsgText <font color="#000080">As</font> String
  <font
color="#000080">If</font> MSN.LocalState = MSTATE_OFFLINE <font color="#000080">Then</font>
    MsgBox &quot;You are not Signed In&quot;
  <font
color="#000080">Else</font>
    <font color="#000080">If</font> MSN.LocalState = MSTATE_INVISIBLE <font
color="#000080">Then</font>
      MsgBox &quot;Change you status first !&quot;
    <font
color="#000080">Else</font>
      <font color="#000080">Set</font> User = MSN.CreateUser(lstOnlineContacts.Text, MSN.Services.PrimaryService)
      bstrMsgText = InputBox(&quot;Enter text to send&quot;, &quot;Send What ?&quot;, &quot;Howdy&quot;, Me.Left, Me.Top)
      User.SendText bstrMsgHeader, bstrMsgText, MMSGTYPE_NO_RESULT
      MsgBox &quot;The following message was sent to &quot; &amp; User.EmailAddress &amp; &quot; : &quot; &amp; bstrMsgText
    <font
color="#000080">End If</font>
  <font color="#000080">End If</font>
<font color="#000080">End Sub</font>
<font
color="#000080">Private Sub</font> Form_Load()
  cmdRefreshList_Click
<font color="#000080">End Sub</font></pre>
  <p>First we fill up the ListBox again but we only fill up the one with the online
  contacts, we don't need to know who's offline only who's online.&nbsp; Then when the user
  presses the Send IM button we make sure that we appear to be online to all users and send
  the message.&nbsp; The code is pretty straightforward.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 6</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="openanimwindow"></a>Open an IM window.</h1>
<table border="0" width="700">
 <tr>
  <td>This is the only example where we will use the MSNAPI variable.&nbsp; Because this
  example is almost exactly the same as the previous one I'll only display the change.
  &nbsp; The first change is the gobal variable declaration:<pre><font color="#000080">Private</font> MSN <font
color="#000080">As New</font> MsgrObject
<font color="#000080">Private</font> MSNAPI <font
color="#000080">As New</font> MessengerAPI.Messenger</pre>
  <p>The second and final change is the cmdSendIM_Click() sub:</p>
  <pre><font color="#000080">Private Sub</font> cmdSendIM_Click()
  MSNAPI.InstantMessage (lstOnlineContacts.Text)
<font
color="#000080">End Sub</font></pre>
  <p>That's it.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 7</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="blockunblockcontacts"></a>Block/unblock contacts.</h1>
<table border="0" width="700">
 <tr>
  <td>To block or unblock a contact we need to do two things.&nbsp; Remove him from one list
  and add him to the other.&nbsp; Create a new project, add the references and on the form
  put 2 Labels (lblBlock &amp; lblAllow), 2 ListBoxes ( lstBlock &amp; lstAllow) and 3
  CommandButtons (cmdAllow, cmdBlock &amp; cmdRefreshLists).&nbsp; And here's the code:<pre><font
color="#000080">Private</font> MSN <font color="#000080">As New</font> MsgrObject
<font
color="#000080">Private Sub</font> RefreshLists()
  lstBlock.Visible = <font color="#000080">False</font>
  lstAllow.Visible = <font
color="#000080">False</font>
  lstBlock.Clear
  lstAllow.Clear
  <font color="#000080">Dim</font> User <font
color="#000080">As</font> IMsgrUser
  <font color="#000080">For Each</font> User <font
color="#000080">In</font> MSN.List(MLIST_BLOCK)
    lstBlock.AddItem (User.EmailAddress)
  <font
color="#000080">Next</font>
  <font color="#000080">For Each</font> User <font
color="#000080">In</font> MSN.List(MLIST_ALLOW)
    lstAllow.AddItem (User.EmailAddress)
  <font
color="#000080">Next</font>
  lstBlock.Visible = <font color="#000080">True</font>
  lstAllow.Visible = <font
color="#000080">True</font>
<font color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdAllow_Click()
  <font
color="#000080">Dim</font> User <font color="#000080">As</font> IMsgrUser
  <font
color="#000080">Set</font> User = MSN.CreateUser(lstBlock.Text, MSN.Services.PrimaryService)
  MSN.List(MLIST_ALLOW).Add User
  MSN.List(MLIST_BLOCK).Remove User
  lstAllow.AddItem (lstBlock.Text)
  lstBlock.RemoveItem (lstBlock.ListIndex)
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdBlock_Click()
  <font
color="#000080">Dim</font> User <font color="#000080">As</font> IMsgrUser
  <font
color="#000080">Set</font> User = MSN.CreateUser(lstAllow.Text, MSN.Services.PrimaryService)
  MSN.List(MLIST_ALLOW).Remove User
  MSN.List(MLIST_BLOCK).Add User
  lstBlock.AddItem (lstAllow.Text)
  lstAllow.RemoveItem (lstAllow.ListIndex)
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdRefreshLists_Click()
  RefreshLists
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> Form_Load()
  RefreshLists
<font
color="#000080">End Sub</font></pre>
  <p>The code isn't really that hard to understand.&nbsp; Make sure you always remember that
  the user you removed from one list has to be added to the other list or your MSN Messenger
  will do weird things when you want to send Instant Messages to that particular person.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 8</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="addremovecontacts"></a>Add/remove contacts.</h1>
<table border="0" width="700">
 <tr>
  <td>Create a new project, add the references and on the form put a Label (lblContacts), a
  ListBox (lstContacts) and 3 CommandButtons (cmdAddContact, cmdRemoveContact &amp;
  cmdRefreshList).&nbsp; Now insert the following code:<pre><font color="#000080">Private</font> MSN <font
color="#000080">As New</font> MsgrObject
<font color="#000080">Private Sub</font> RefreshList()
  lstContacts.Visible = <font
color="#000080">False</font>
  lstContacts.Clear
  <font color="#000080">Dim</font> User <font
color="#000080">As</font> IMsgrUser
  <font color="#000080">For Each</font> User <font
color="#000080">In</font> MSN.List(MLIST_CONTACT)
    lstContacts.AddItem (User.EmailAddress)
  <font
color="#000080">Next</font>
  lstContacts.Visible = <font color="#000080">True</font>
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdAddContact_Click()
  <font
color="#000080">Dim</font> User <font color="#000080">As</font> IMsgrUser
  <font
color="#000080">Set</font> User = MSN.CreateUser(InputBox(&quot;Enter users e-mail address.&quot;, &quot;Add Contact&quot;, &quot;&quot;, Me.Left, Me.Top), _
  MSN.Services.PrimaryService)
  MSN.List(MLIST_CONTACT).Add User
  MSN.List(MLIST_ALLOW).Add User
  lstContacts.AddItem User.EmailAddress
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdRefreshList_Click()
  RefreshList
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> cmdRemoveContact_Click()
<font
color="#000080">On Error Resume Next</font>
  <font color="#000080">Dim</font> User <font
color="#000080">As</font> IMsgrUser
  <font color="#000080">Set</font> User = MSN.CreateUser(lstContacts.Text, MSN.Services.PrimaryService)
  MSN.List(MLIST_ALLOW).Remove User
  MSN.List(MLIST_CONTACT).Remove User
  MSN.List(MLIST_BLOCK).Remove User
  lstContacts.RemoveItem (lstContacts.ListIndex)
<font
color="#000080">End Sub</font>
<font color="#000080">Private Sub</font> Form_Load()
  RefreshList
<font
color="#000080">End Sub</font></pre>
  <p>First we add all of our contacts to the ListBox.&nbsp; Then when a user wants to add a
  contact they get an InputBox to enter the user's e-mail address into.&nbsp; We set the
  User variable to that e-mail address and add it to the allow and contact list.&nbsp; Then
  we add it to the ListBox.&nbsp;&nbsp; When removing a contact we must make sure we remove
  the User object from whatever list it might be on, we don't know if that particular
  contact is blocked or not.&nbsp; Don't forget the error handling as we will get an error
  because the contact is either on the allow or block list, not both.&nbsp; Then finally we
  remove the user from the ListBox.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 9</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="handlingmsnevents"></a>Handling MSN events.</h1>
<table border="0" width="700">
 <tr>
  <td>This particular topic doesn't have a project to go along with it.&nbsp; I'm going to
  take you thru the most important MSN events and explain how to use them.&nbsp; The
  following topic is an example of how to use the MSN events in your program.<pre><font
color="#000080">Private Sub</font> msn_OnListAddResult(<font color="#000080">ByVal</font> hr <font
color="#000080">As Long</font>, <font color="#000080">ByVal</font> MLIST <font color="#000080">As</font> Messenger.MLIST, <font
color="#000080">ByVal</font> pUser <font color="#000080">As</font> _
  Messenger.IMsgrUser)
<font
color="#000080">End Sub</font></pre>
  <p>This will happen when a user is added to any one of the 4 lists.&nbsp; This is
  important so you can update your contacts list, when a user is blocked he or she will be
  ADDED to the block list.&nbsp; pUser tells you which user was added to the list and MLIST
  specifies which list the user was added to.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnListRemoveResult(<font color="#000080">ByVal</font> hr <font
color="#000080">As Long</font>, <font color="#000080">ByVal</font> MLIST <font color="#000080">As</font> Messenger.MLIST, <font
color="#000080">ByVal</font> pUser <font color="#000080">As</font> _
  Messenger.IMsgrUser)
<font
color="#000080">End Sub</font></pre>
  <p>This will happen when a user is removed from any one of the 4 lists.&nbsp; This is also
  important so you can update your contacts list, when a user is unblocked he or she will be
  REMOVED from the block list.&nbsp; pUser tells you which user was removed from the list
  and MLIST specifies which list the user was removed from.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnLocalFriendlyNameChangeResult(<font
color="#000080">ByVal</font> hr <font color="#000080">As Long</font>, <font color="#000080">ByVal</font> pService <font
color="#000080">As</font> _
  Messenger.IMsgrService, <font color="#000080">ByVal</font> bstrPrevFriendlyName <font
color="#000080">As String</font>)
<font color="#000080">End Sub</font></pre>
  <p>This will happen when your nickname has been changed.&nbsp; pService specifies the
  service you are currently connected to and bstrPrevFriendlyName tells you what your
  nickname was before it was changed.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnLocalStateChangeResult(<font
color="#000080">ByVal</font> hr <font color="#000080">As Long</font>, <font color="#000080">ByVal</font> mLocalState <font
color="#000080">As</font> _
  Messenger.MSTATE, <font color="#000080">ByVal</font> pService <font
color="#000080">As</font> Messenger.IMsgrService)
<font color="#000080">End Sub</font></pre>
  <p>This will happen when your state has been changed.&nbsp; mLocalState tells you what
  your current state is and pService specifies the service you are connected to.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnLogoff()
<font color="#000080">End Sub</font></pre>
  <p>I don't think I need to explain this one.&nbsp; Just keep in mind that if the user
  exits MSN that this event will not be triggered, only if he or she shuts down his or her
  computer or if they just Sign Out.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnLogonResult(<font color="#000080">ByVal</font> hr <font
color="#000080">As Long</font>, <font color="#000080">ByVal</font> pService <font
color="#000080">As</font> Messenger.IMsgrService)
<font color="#000080">End Sub</font></pre>
  <p>This will happen when the user has successfully Signed In.&nbsp; pService tells you
  which service the user Signed In to.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnTextReceived(<font color="#000080">ByVal</font> pIMSession <font
color="#000080">As</font> Messenger.IMsgrIMSession, <font color="#000080">ByVal</font> _
  pSourceUser <font
color="#000080">As</font> Messenger.IMsgrUser, <font color="#000080">ByVal</font> bstrMsgHeader <font
color="#000080">As String</font>, <font color="#000080">ByVal</font> bstrMsgText <font
color="#000080">As</font> _
  <font color="#000080">String</font>, pfEnableDefault <font
color="#000080">As Boolean</font>)
<font color="#000080">End Sub</font></pre>
  <p>This will happen when the user receives text.&nbsp; This could mean 2 things :</p>
  <p>1) Someone has said something to the user.</p>
  <p>2) Someone has started to type something to the user but has not yet sent it.</p>
  <p>pIMSession tells you which session the event has been sent to, pSourceUser tells you
  which user sent the text, bstrMsgHeader contains information regarding the font of the
  text and if it is the second case (see above) this is something else, bstrMsgText lets you
  know which text was sent and pfEnableDefault is used to block the information from going
  to the IM box.&nbsp; If you set pfEnableDefault to False then the user will not receive
  the message.µ</p>
  <pre><font color="#000080">Private Sub</font> msn_OnUnreadEmailChanged(<font
color="#000080">ByVal</font> MFOLDER <font color="#000080">As</font> Messenger.MFOLDER, <font
color="#000080">ByVal</font> cUnreadEmail _
  <font color="#000080">As Long</font>, pfEnableDefault <font
color="#000080">As Boolean</font>)
<font color="#000080">End Sub</font></pre>
  <p>This will happen when the user receives an e-mail.&nbsp; MFOLDER tells you which folder
  it was sent to (Inbox or not), cUnreadEmail is the amount of unread e-mails in total and
  if you set pfEnableDefault to False the user won't get that popup telling them.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnUserFriendlyNameChangeResult(<font
color="#000080">ByVal</font> hr <font color="#000080">As Long</font>, <font color="#000080">ByVal</font> pUser <font
color="#000080">As</font> _
  Messenger.IMsgrUser, <font color="#000080">ByVal</font> bstrPrevFriendlyName <font
color="#000080">As String</font>)
<font color="#000080">End Sub</font></pre>
  <p>This will happen when a user changes his or her nickname.&nbsp; pUser tells you which
  user changed their nickname and bstrPrevFriendlyName tells you what their nickname was
  before they changed it.&nbsp; TIP : to find out what their current nickname is use this :
  pUser.FriendlyName.</p>
  <pre><font color="#000080">Private Sub</font> msn_OnUserStateChanged(<font color="#000080">ByVal</font> pUser <font
color="#000080">As</font> Messenger.IMsgrUser, <font color="#000080">ByVal</font> mPrevState <font
color="#000080">As</font> _
  Messenger.MSTATE, pfEnableDefault <font color="#000080">As Boolean</font>)
<font
color="#000080">End Sub</font></pre>
  <p>This will happen when a user changes his or her status.&nbsp; pUser tells you which
  user changed their status and mPrevState tells you what their status was before they
  changed it.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
<h1><a name="creatingaselfupdatingcontactslist"></a>Creating a self-updating contacts
list.</h1>
<table border="0" width="700">
 <tr>
  <td>To do this we're just going to make a few minor adjustments to a previous project, the
  one that <a href="#sendyourcontactslisttoatreeview">sends your contact list to a treeview</a>.
  &nbsp; The first change we'll make is the variable declaration, this is what it should
  look like now:<pre><font color="#000080">Private WithEvents</font> MSN <font
color="#000080">As</font> MsgrObject</pre>
  <p>The second change is the Form_Load() sub, it should look like this:</p>
  <pre><font color="#000080">Private Sub</font> Form_Load()
  <font color="#000080">Set</font> MSN = <font
color="#000080">New</font> MsgrObject
  RefreshList
<font color="#000080">End Sub</font></pre>
  <p>Then finally we have to add the following:</p>
  <pre><font color="#000080">Private Sub</font> MSN_OnUserStateChanged(<font color="#000080">ByVal</font> pUser <font
color="#000080">As</font> Messenger.IMsgrUser, <font color="#000080">ByVal</font> mPrevState <font
color="#000080">As</font> _
  Messenger.MSTATE, pfEnableDefault <font color="#000080">As Boolean</font>)
  RefreshList
<font
color="#000080">End Sub</font></pre>
  <p>That's it.</p>
  <p><font color="#FF0000">The code for this example can be found in folder number 10</font>.</p>
  <p><font color="#000000"><a href="#top">Back to top</a></font></td>
 </tr>
</table>
</body>
</html>

