<div align="center">

## VBScript Tutorial: Chapter 2\-\-Introduction to the IE Object Model


</div>

### Description

Since you are using VBScript, that should mean that either you have separate versions for Netscape and Internet Explorer users, make pages which tailor themselves to the client's browser, use VBScript only for non-essential page features, or are developing only for people using Internet Explorer. I have provided an example of a forwarding page which separates users capable of VBScript from others, and also a script which calls extra features if the client is using Internet Explorer. Note that they both use JavaScript, so that they will function in either NN of IE.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David M\. Carr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-m-carr.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-m-carr-vbscript-tutorial-chapter-2-introduction-to-the-ie-object-model__4-6190/archive/master.zip)





### Source Code

```
If you can fit into one of these four categories, you can use VBScript and other IE-specific features, such as floating frames, ActiveX Controls, IE-Dynamic HTML, etc.  In this tutorial, I assume that everyone is using Internet Explorer to view it, and thus I use floating frames whenever they are the convenient way to accomplish my goal.  If you only want to make one version of your site, using only one scripting language, then you will probably be best off using JavaScript as your primary scripting language.Now, some pieces of code are built-in to VBScript. Some types of code which are always available to VBScript are statements, operators, and built-in functions. In addition to these, VBScript in web pages also has access to the browser's object model.  This is a structured way of looking at how the browser works. It gives access to properties, events and methods. Events let you know that something has happened, such as a mouse click, and let you respond. Properties store information, such as background color, and often can be changed. Properties which cannot be changed are called read-only properties. Methods are the things the object can do.  Examples of methods are Alert, which displays a message, and Write, which writes information to a document.Internet Explorer 3 and greater support VBScript, but each version has a different object model. IE4's object model is much more extensive than IE3's. It is very tempting to ignore IE3, and use all of the capabilities of the latest version, but generally it is best to write for the greatest amount of the audience as is possible, which often includes users with older browsers.Here's a simplified look at the IE3 Object Model. (IE4's Object Model is much too complex to cover on this page)
Window
Navigator
Location
History
Frames...
Document
Links...
Anchors...
Forms...
Elements...
If those ellipses (...) look suspicious to you, it's for good reason. They indicate that there may be more than one of that object. Each frame is treated as a window object, and may contain other frames. Objects are referred to using dot notation, separating the objects by a dot. For example, to refer to the Alert method, you could use Window.Alert. If you wished to refer to the background color property of the document, you could use Window.Document.BgColor. In most cases, the window object is left out. It is understood by the browser if it is omitted, because Window is the top-level object. Thus, the previous examples could also be written Alert and Document.BgColor.A more in-depth look at the IE Object Model is given in the Object Model Reference.
Also, you can see what functions are built into VBScript in the MS VBScript Reference.
```

