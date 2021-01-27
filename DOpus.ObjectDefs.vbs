'* Int        - a whole number
'* Currency   - this is a standard variant type, and is used by Opus in just a couple of cases that require numbers larger than an Int can hold (unfortunately ActiveX scripting does not support a 64 bit integer type).
'* String     - a text string
'* Bool       - a Boolean (either True or False)
'* Date       - a date/time value
'* Collection - a collection of multiple objects of (generally) the same type. Collections can be easy enumerated in some languages (e.g. in VBScript, using the For Each construct). Collections are only returned by Opus - unlike an array (or a Vector object), they are never created or modified directly (although some script methods can be used to modify them in certain cases). For example, the Tab object has a property called selected that represents all currently selected items in that tab - it is a collection of items.
'* Object     - an Opus script object with defined methods and properties (one of the object types listed below). Some objects can be both an object and a collection - that is, they have methods and properties, but can also be enumerated like a collection. Some objects also have a default value - that is, simply using the object's name without any method or property will return a value of its own. For example, the Metadata object's default value is a string indicating the primary type of metadata available. You can also refer to an object's default value using the special def_value property name.
'* Variant    - a variable of any type (this is used with a few objects - for example, a Var object can store a variant)

' ! Red highlighted comment
' ? Blue highlighted comment
' * Green highlighted comment
' todo orange highlighted comment
' // Grey with strikethrough comment


' The DOpus object is one of the two global script objects provided by Opus, and is available to all scripts. It provides various helper methods, and collections that let you access things like Listers and toolbars.
Class DOpus
    ' The Aliases object gives the script access to the defined folder aliases.
    Property Get Aliases ' Return Type object:Aliases
    End Property

    ' Returns a collection of Format objects representing the used-defined favorite formats.
    Property Get FavoriteFormats ' Return Type collection:Format 
    End Property

    ' Returns a Favorites object which lets you query and modify the user-defined favorite folders.
    Property Get Favorites ' Return Type object:Favorites
    End Property

    ' Returns a FiletypeGroups object which lets you enumerate and query the configured file type groups.
    Property Get FileTypeGroups ' Return Type object:FiletypeGroups 
    End Property

    ' Returns a GlobalFilters object which lets you access information about the global filter settings (configured on the Folders / Global Filters page in Preferences).
    Property Get Filters ' Return Type object:GlobalFilters 
    End Property

    ' Returns a string representing the current user interface language.
    Property Get Language ' Return Type string
    End Property

    ' Returns a Listers object which represents any currently open Lister windows (each one is represented by a Lister object).
    Property Get Listers ' Return Type object:Listers 
    End Property

    ' Returns a SmartFavorites object which lets you query the SmartFavorites data.
    Property Get SmartFavorites ' Return Type object:SmartFavorites
    End Property

    ' Returns a ScriptStrings object which lets your script access any strings defined as string resources.
    Property Get Strings ' Return Type object:ScriptStrings 
    End Property

    ' Returns a TabGroups object which lets your script access and manipulate the configured folder tab groups.
    Property Get TabGroups ' Return Type object:TabGroups 
    End Property

    ' This Vars object represents all defined variables with global scope.
    Property Get Vars ' Return Type object:Vars 
    End Property

    ' The Version object provides information about the current Opus program version.
    Property Get Version ' Return Type object:Version 
    End Property

    ' Returns a Viewers object which represents any currently open standalone image viewers (each one is represented by a Viewer object).
    Property Get Viewers ' Return Type object:Viewers 
    End Property

    ' Clears the script output log.
    Sub ClearOutput
    End Sub

    ' Creates and returns a new DOpusFactory object, which can be used to create various lightweight helper objects like Blob, Map and Vector.
    Function Create ' Return Type object:DOpusFactory 
    End Function

    ' Delays for the specified number of milliseconds before returning.
    Sub Delay(<int:time>)
    End Sub

    ' Creates a new Dialog object, that lets you display dialogs and popup menus.
    Function Dlg ' Return Type object:Dialog
    End Function

    ' Creates the DPI helper object which assists when dealing with different system scaling settings (e.g. high-DPI monitors).
    Function DPI ' Return Type object:DPI 
    End Function

    ' Creates a new FSUtil object, that provides helper methods for accessing the file system.
    Function FSUtil ' Return Type object:FSUtil 
    End Function

    ' Retrieves the current contents of the system clipboard, if it contains either text or files.
    Function GetClip(none or <string:type>) ' Return Type string
    End Function

    ' Returns a string indicating the native format of the clipboard contents - ""text"", ""files"" or an empty string in any other case.
    Function GetClipFormat ' Return Type string
    End Function

    ' Returns a string indicating which qualifier keys are currently held down. If none are held down, the string will be ""none"". Otherwise, the string can contain any or all of the following, separated by commas: ""shift"", ""ctrl"", ""alt"", ""lwin"", ""rwin"".
    Function GetQualifiers ' Return Type string
    End Function

    ' Loads an image file from the specified external file. You can optionally specify the desired size to load the image at, and whether the alpha channel (if any) should be loaded or not.
    Function LoadImage(<string:filename>,[<int:width>],[<int:height>],[<bool:alpha>]) ' Return Type object:Image 
    End Function

    ' Extracts a thumbnail from the specified external file. You can optionally specify a timeout (in milliseconds) and the desired size to load the thumbnail at.
    Function LoadThumbnail(<string:filename>,[<int:timeout>],[<int:width>],[<int:height>]) ' Return Type object:Image or bool (False)
    End Function

    ' Prints the specified text string to the script output log (found in the Utility Panel,  the CLI in script mode, the Rename dialog and the Command Editor in script mode).
    Sub Output(<string:text>,[<bool:error>],[<bool:timestamp>])
    End Sub

    ' Causes Opus to reload and reinitialize the specified script. You must provide the full pathname of the script on disk (if a script add-in wants to reload itself you can pass the value of the Script.file property).
    Sub ReloadScript(<string:file>)
    End Sub

    ' Places the specified text, or Item collection (or Vector of Item objects) on the system clipboard. If called with no arguments the clipboard will be cleared.
    Sub SetClip(<string:text> or collection:Item or none)
    End Sub

    ' Returns a Toolbars object which lets you enumerate all defined toolbars (whether they are currently open or not).
    Function Toolbars(<string:type>) ' Return Type object:Toolbars 
    End Function

    ' Returns a string indicating the type of an object or variable.
    Function TypeOf(any) ' Return Type string
    End Function
End Class

'The Script object is one of the two global script objects provided by Opus. This object is provided to script addins when their various event handlers are invoked (other than for the OnInit event). It provides information relating to the script itself.
Class Script
    ' Returns a ScriptConfig object representing the configuration values for this script. In the OnInit method a script can define the properties that make up its configuration - the user can then edit these values in Preferences. The object returned by the config property represents the values that the user has chosen.
    Property Get config ' Return Type object:ScriptConfig 
    End Property

    ' Returns the path and filename of this script.
    Property Get file ' Return Type string
    End Property

    ' Returns a Vars object that represents the variables that are scoped to this particular script. This allows scripts to use variables that persist from one invocation of the script to another.
    Property Get vars ' Return Type object:Vars 
    End Property

    ' Returns True if local HTTP help is enabled (that is, if help is shown in the user's web browser), False if the old HtmlHelp-style help is enabled. If HTTP help is enabled, your script is able to add its own help pages via the OnGetHelpContent event, and it can trigger the display of its own help pages using the ShowHelp method.
    Function HttpHelpEnabled ' Return Type bool
    End Function

    ' If your script implements the OnAddColumns event, you can call the InitColumns method at any time to reinitialize your columns. You may want to do this, for example, in response to the user modifying your script's configuration.
    Sub InitColumns
    End Sub

    ' If your script implements the OnAddCommands event, you can call the InitCommands method at any time to reinitialize your commands. You may want to do this, for example, in response to the user modifying your script's configuration.
    Sub InitCommands
    End Sub

    ' Using the OnGetHelpContent event your script can add its own content to the F1 help. If your script is bundled as a script package you can include .html files in a sub-directory of the package called help, and then load them easily using this method. You can then pass the loaded data to the GetHelpContentData.AddHelpPage method to add the page.
    Function LoadHelpFile(<string:name>) ' Return Type string
    End Function

    ' If your script is bundled as a script package you can include PNG and JPG image files in a sub-directory of the package called help, and then load them easily using this method. You can then pass the loaded data to the GetHelpContentData.AddHelpImage method to add the image.
    Function LoadHelpImage(<string:name>) ' Return Type object:Blob 
    End Function

    ' Loads an image file from the specified external file. If your script is bundled as a script package you can place image files in a sub-directory of the package called images and then load them from your script by giving their name. You can optionally specify the desired size to load the image at, and whether the alpha channel (if any) should be loaded or not.
    Function LoadImage(<string:name>,[<int:width>],[<int:height>],[<bool:alpha>]) ' Return Type object:Image 
    End Function

    ' Loads external script resources and makes them available to the script. You can either provide a filename or a raw XML string. If your script is bundled as a script package, the resource file must have a .odxml extension for LoadResources to be able to find it in the package.
    Sub LoadResources(<string:name> or <string:XML>)
    End Sub

    ' If your script implements any custom columns, you can use this method to cause them to be regenerated if they are currently shown in any tabs. You may want to do this, for example, in response to the user modifying your script's configuration. Pass the name of the column you want to regenerate as the argument to this method.
    Sub RefreshColumn(<string:name>)
    End Sub

    ' If your script adds its own help pages via the OnGetHelpContent event, and the user has http help enabled, you can call this method to display your help in the user's web browser. You might want to do this when the user clicks a Help button in your script dialog, for example. You can use the HttpHelpEnabled method to check if http help is enabled before calling this function.
    Sub ShowHelp(<string:page>)
    End Sub
End Class



' There are also a number of objects that Opus provides as parameters to methods within a script. For example, when a script function is invoked (e.g. when the button is clicked), Opus calls its OnClick method, passing it a ClickData object.


'This object represents a folder alias, and is retrieved using the Aliases object.
Class Alias 
    ' 'Returns the name of the alias.
    ' Default Property Get Def_value 'Return Type string
    ' End Property

    'Returns the target of the alias as a Path object.
    Property Get path ' Return Type object:Path 
    End Property

    'True if the object is a system-defined alias, False if it is user defined.
    Property Get system ' Return Type bool
    End Property
End Class

'This object is a collection of all defined folder aliases. It is retrieved using the DOpus.aliases collection property.
Class Aliases ' Default Return Type collection:Alias 
    'Adds a new alias to the system with the specified name and path. Note that you should not provide the leading forward-slash (/) in the alias name.
    Sub Add(<string:name>,<string:path>)
    End Sub

    'Deletes the specified alias.
    Sub Delete
    End Sub

    'Updates the state of this object. When the Aliases object is first retrieved via DOpus.aliases, a snapshot is taken of the aliases at that time. If you make changes via the object it will reflect them but any changes made outside the script (e.g. via the Favorites ADD=alias command) will not be detected unless you call the Update method.
    Sub Update
    End Sub
End Class

'This object represents the arguments supplied on the command line for script-defined internal commands. It is retrieved from the ScriptCommandData.Func.args property.
Class Args
    'The Args object will have one property corresponding to each of the arguments in the command line template.
    Property Get argument_name ' Return Type variant
    End Property

    'The got_arg property returns an object with a bool child property for each argument in the template. It lets you test if a particular argument was provided on the command line, before you actually query for the value of the argument. For example, If Args.got_arg.size Then...
    Property Get got_arg ' Return Type object
    End Property
End Class

'This object provides metadata properties relating to audio files. It is obtained from the Metadata object.
Class AudioMeta
    'Returns the value of the specified column, as listed in the Music section of the Keywords for Columns page.
    Property Get column_keyword ' Return Type variant
    End Property

    'Returns a collection of AudioCoverArt objects representing any cover art imagery stored in the audio file.
    Property Get coverart ' Return Type collection:AudioCoverArt 
    End Property
End Class

'This object provides access to an audio file's embedded cover art. It is obtained from the AudioMeta.coverart property.
Class AudioCoverArt ' Default Return Type string
    'Returns a Blob object representing the actual image data.
    Property Get data ' Return Type object:Blob 
    End Property

    'Returns the bit depth of this image.
    Property Get depth ' Return Type int
    End Property

    'Returns the description of this image (if any).
    Property Get desc ' Return Type string
    End Property

    'Returns the default file extension for this image, if it can be determined.
    Property Get ext  ' Return Type string
    End Property

    'Returns the height of this image, in pixels.
    Property Get height ' Return Type int
    End Property

    'Returns the image's MIME type, if specified in the file.
    Property Get mime ' Return Type string
    End Property

    'Returns a FileSize object representing the size of the image data.
    Property Get size ' Return Type object:FileSize 
    End Property

    'Returns a ""pretty"" form of the intended use string (i.e. the default value), translated to the current Opus user interface language.
    Property Get type ' Return Type string
    End Property

    'Returns the width of this image, in pixels.
    Property Get width ' Return Type int
    End Property
End Class

'This object provides a simple interface for dealing with binary data. It is obtained from the DOpusFactory.Blob method and also returned by the AudioCoverArt.data property.
Class Blob
    'Returns a FileSize object representing the size of this Blob in bytes.
    Property Get size ' Return Type object:FileSize 
    End Property

    'Compares the contents of this Blob against another Blob (or array). By default the entire contents of the two blobs are compared. The optional parameters that let you configure the operation are:
    Function Compare(<Blob:source>,<int:to>,<int:from>,<int:size>) ' Return Type int
    End Function

    'Copies data from the source Blob (or array) into this Blob. By default the entire contents of the source Blob will be copied over the top of this one. The optional parameters that let you configure the operation are:
    Sub CopyFrom(<Blob:source>,<int:to>,<int:from>,<int:size>)
    End Sub

    Sub CopyFrom(<string>,<type>)
    End Sub

    'Searches the contents of this Blob for the data contained in another Blob (or array). By default the entire contents of this Blob are searched. The optional from parameter lets you specify the starting position for the search, and the optional size parameter lets you specify the length of data in this Blob to search through.
    Function Find(<Blob:search>,<int:from>,<int:size>) ' Return Type object:FileSize 
    End Function

    'Frees the memory associated with this Blob and resets its size to 0.
    Sub Free
    End Sub

    'Initialises the contents of the Blob (every byte within the blob will be set to 0). Equivalent to Set(0).
    Sub Init
    End Sub

    'Resizes the Blob to the specified number of bytes.
    Sub Resize
    End Sub

    'Reverses the contents of the Blob.
    Sub Reverse
    End Sub

    'Sets the contents of the Blob to the specified byte value (every byte within the blob will be set to that value). By default the whole Blob will be affected. The option to parameter lets you specify a byte offset to start at, and the optional size parameter lets you control the number of bytes affected.
    Sub Set(<byte:value>,<int:to>,<int:size>)
    End Sub

    'Converts the contents of this Blob to a SAFEARRAY of type VT_UI1. By default the entire contents of the Blob will be copied to the array. The optional parameters that let you configure the operation are:
    Function ToArray(<int:from>,<int:size>) ' Return Type SAFEARRAY of VT_UI1
    End Function

    'Converts the contents of this Blob to a SAFEARRAY of type VT_VARIANT. Each variant in the array contains a VT_UI1. By default the entire contents of the Blob will be copied to the array. The optional parameters that let you configure the operation are:
    Function ToVBArray ' Return Type SAFEARRAY of VT_VARIANT
    End Function

End Class

'A BusyIndicator object lets you control the breadcrumbs bar busy indicator from your script.
Class BusyIndicator

    'Before the Init method has been called, you can set this property to True to enable abort by the user (as shown above).
    Property Get abort ' Return Type bool
    End Property

    ' Removes the busy indicator from display and destroys its internal data structures. The BusyIndicator object itself can be re-used by calling the Init method again.
    Sub Destroy
    End Sub
    ' Removes the busy indicator from display, but does not destroy its internal data. The indicator can be re-displayed by calling the Show method.
    Sub Hide

    End Sub

    ' Initializes a BusyIndicator object and optionally displays it. The window parameter specifies the Lister that the indicator is to be attached to - you can pass either a Lister or a Tab object.
    Function Init(<object:window>,<string:description>,<bool:visible>) ' Return Type bool
    End Function

    ' Displays the busy indicator.
    Sub Show
    End Sub

    ' Updates the busy indicator. The description parameter lets you specify a new description string, and the optional percentage parameter lets you specify a new percentage complete value from 0 to 100.
    Sub Update(<string:description>,<int:percentage>)
    End Sub

End Class

'This object represents a column that has been added to the display in a tab. A collection of columns can be obtained from the Format object.
Class Column ' Default Return name of the column as string
    'Returns True if the column width is set to auto.
    Property Get autosize ' Return Type bool
    End Property

    'Returns True if the column width is set to collapse.
    Property Get collapse ' Return Type bool
    End Property

    'Returns True if the column width is set to expand.
    Property Get expand ' Return Type bool
    End Property

    'Returns True if the column width is set to fill.
    Property Get fill ' Return Type bool
    End Property

    'Returns the name of the column as displayed in the Lister column header.
    Property Get header ' Return Type string
    End Property

    'Returns the name of the column as displayed in the Columns tab in the Folder Options dialog.
    Property Get label ' Return Type string
    End Property

    'Returns the maximum width of the column in pixels, or the string ""fill"" if the maximum is set to fill.
    Property Get max ' Return Type int or string
    End Property

    'Returns the name of the column.
    Property Get name ' Return Type string
    End Property

    'Returns True if the sort direction of the column is reversed.
    Property Get reverse ' Return Type bool
    End Property

    'Returns the sort order of the column (e.g. 1 for the primary sort field, 2 for the secondary sort field, etc). Returns 0 if the display is not sorted by this column.
    Property Get sort ' Return Type int
    End Property

    'Returns the current display width of the column in pixels.
    Property Get width ' Return Type int
    End Property
End Class

'This object is used to run Opus commands. It is obtained from the ScriptCommandData.func or ClickData.func properties, and can also be created by the DOpusFactory.Command method.
Class Command


    'Set this property to False to prevent files used by this command from being deselected, and True to deselect them once the function is finished. Note that files will only be deselected if they came from a Tab object, and only then if the command is successful.
    Property Get deselect ' Return Type bool
    End Property

    'Returns a Path object that represents the destination folder of this command. If a destination tab is set, this will be the path in the tab. You can not set this property directly - instead, use either the SetDest or SetDestTab methods to change the destination folder.
    Property Get dest ' Return Type object:Path 
    End Property

    'Returns a Tab object that represents the destination tab for this command (if it has one - not all commands require a destination). You can not set this property directly - instead, use the SetDestTab method to change the destination tab.
    Property Get desttab ' Return Type object:Tab 
    End Property

    'Returns the number of items in the files collection.
    Property Get filecount ' Return Type int
    End Property

    'Returns a collection of all Item objects that represent the files and folders this command is to act upon. You can not modify this collection directly - instead you can use the various methods (ClearFiles, SetFiles, AddFile, RemoveFile, etc.) to modify the list of items to act upon.
    Property Get files ' Return Type collection:Item 
    End Property

    'Returns the number of instruction lines added to the command.
    Property Get linecount ' Return Type int
    End Property

    'Returns a Progress object that you can use to display a progress indicator to the user.
    Property Get progress ' Return Type object:Progress 
    End Property

    'After every command that is run with this object, a Results object is available from this property. This provides information about the outcome of the command.
    Property Get results ' Return Type object:Results 
    End Property

    'Returns a Path object that represents the source folder of this command. If a source tab is set, this will be the path in the tab. You can not set this property directly - instead, use either the SetSource or SetSourceTab methods to change the source folder.
    Property Get source ' Return Type object:Path 
    End Property

    'Returns a Tab object that represents the source tab for this command. You can not set this property directly - instead, use the SetSourceTab method to change the source tab.
    Property Get sourcetab ' Return Type object:Tab 
    End Property

    'This Vars object represents all defined variables with command scope (that are scoped to this function - e.g. that were set using the @set directive).
    Property Get vars ' Return Type object:Vars 
    End Property

    ' Adds the specified item to the collection of items this command is to act upon. You can pass the item's path as either a string or a Path object, and you can also pass an Item object directly.
    Function AddFile(<string:path> or <Path:path> or <Item:item>) ' Return Type int
    End Function

    ' Adds the items in the specified collection to the list of items this command is to act upon. The return value is the new number of items in the collection.
    Function AddFiles(collection:Item or Vector:Item or Vector:Path or Vector:string) ' Return Type int
    End Function

    ' Adds the contents of the clipboard to the collection of items this command is to act upon. This method supports both files and file paths copied to the clipboard as text. The return value is the new number of items in the collection.
    Function AddFilesFromClipboard ' Return Type int
    End Function

    ' Reads file paths from the contents of the specified file and adds them to the item collection. You can provide the file's path as either a string or a Path object. The file must consist of one absolute path per line.
    Function AddFilesFromFile(<string:path>,<string:encoding>) ' Return Type int
    End Function

    ' Adds the contents of the specified folder to the collection of items this command is to act upon. You can pass the folder's path as either a string or a Path object. You can also append a wildcard pattern to the path to only add files matching the specified pattern.
    Function AddFilesFromFolder ' Return Type int
    End Function

    ' Adds the specified instruction line to the command that this object will run. The AddLine method lets you build up complicated multiple line commands - add each line in turn and then run the command using the Run method. For a single line command it is simpler to use the RunCommand method.
    Sub AddLine
    End Sub

    ' Clears all instruction lines from the command.
    Sub Clear
    End Sub

    ' Clears the failure flags from the Item collection. Any items that fail when a command is run will have their failed property set to True, and once this has happened the file will be skipped over by any subsequent commands. You can call this method to reset all the failure flags.
    Sub ClearFailed
    End Sub

    ' Clears the collection of items this command is to act upon.
    Sub ClearFiles
    End Sub

    ' Clears any modifiers that have been set for this command. The supported modifiers are a subset of the full list of command modifiers - see the SetModifier method for a list of these. You can also pass * to clear all modifiers that have been set.
    Sub ClearModifier
    End Sub

    ' Returns a StringSet containing the names of all the Opus commands. You can optionally filter this set by providing one or more of the following flags as an argument to the CommandList method:
    Function CommandList(none or <string:types>) ' Return Type object:StringSet 
    End Function

    ' Creates a new Dialog object, that lets you display dialogs and popup menus. The dialog's window property will be automatically assigned to the source tab.
    Function Dlg ' Return Type object:Dialog
    End Function

    ' Returns a Map of the modifiers that have been set for this command (either by the SetModifier method, or in the case of script add-ins any modifiers that were set on the button that invoked the script).
    Function GetModifiers ' Return Type object:Map 
    End Function

    ' Returns True if the specified Set command condition is true. This is the equivalent of the @ifset command modifiers. The optional second parameter lets you test a condition based on a command other than Set - for example, IsSet(""VIEWERCMD=mark"", ""Show"") in the viewer to test if the current image is marked.
    Function IsSet(<string:condition>,[<string:command>]) ' Return Type bool
    End Function

    ' Removes the specified file from the Item collection. You can pass the file's path as either a string or a Path object. You can also pass the Item itself, or its index (starting from 0) within the collection. The return value is the new number of items in the collection.
    Function RemoveFile(<string:path> or <Path:path> or <Item:item> or <int:index>) ' Return Type int
    End Function

    ' Runs the command that has been built up with this object. The return value indicates whether or not the command ran successfully. Zero indicates the command could not be run or was aborted; any other number indicates the command was run for at least some files. (Note that this is not the ""exit code"" for external commands. For external commands it only indicates whether or not Opus launched the command. If you need the exit code of an external command, use the WScript.Shell Run or Exec methods to run the command.) You can use the Results property to find out more information about the results of the command, and also discover which files (if any) failed using the failed property of each Item in the files collection.
    Function Run ' Return Type int
    End Function

    ' Runs the single line command given by the instruction argument. Calling this method is the equivalent of adding the single line with the AddLine method and then calling the Run method.
    Function RunCommand ' Return Type int
    End Function

    ' Sets the command's destination to the specified path. You can provide the path as either a string or a Path object. Calling this method clears the destination tab property from the command.
    Sub SetDest
    End Sub

    ' Sets the command's destination to the specified tab. The destination path will be initialized from the tab automatically (so you don't need to call SetDest as well as SetDestTab).
    Sub SetDestTab
    End Sub

    ' Configures the command to use the files in the specified Item collection as the items the command will act upon.
    Sub SetFiles
    End Sub

    ' Turns on a modifier for this command. The supported modifiers are a subset of the full list of command modifiers:
    Sub SetModifier(<string:modifier>,<string:value>)
    End Sub

    ' Lets you share the progress indicator from one command with another command. You can pass this method the value of progress property obtained from another Command object.
    Sub SetProgress
    End Sub

    ' This method lets you control which qualifier keys the command run by this object will consider to have been pressed when it was invoked. For example, several internal commands change their behavior when certain qualifier keys are held down - calling this method allows you to set which keys they will see.
    Sub SetQualifiers
    End Sub

    ' Sets the command's source to the specified path. You can provide the path as either a string or a Path object. Calling this method clears the source tab property from the command.
    Sub SetSource
    End Sub

    ' Sets the command's source to the specified tab. The source path will be initialized from the tab automatically (so you don't need to call SetSource as well as SetSourceTab).
    Sub SetSourceTab
    End Sub

    ' Sets the type of function that this command will run. This is equivalent to the drop-down control in the Advanced Command Editor. The type argument must be one of the following strings: std, msdos, script, wsl. Standard (std) is the default if the type is not specifically set.
    Sub SetType
    End Sub

    ' This method can be used to update the appearance of toolbar buttons that use @toggle:if to set their selection state based on the existence of a global-, tab- or Lister-scoped variable. You would call this method if you have changed such a variable from a script to force buttons that use it to update their selection status.
    Sub UpdateToggle
    End Sub

End Class

'The Control object represents a control on a script dialog; it lets you read and modify a control's value (and contents).
Class Control
    ' Set or query the color used for the background (fill) of this control. This is in the format #RRGGBB (hexadecimal) or RRR,GGG,BBB (decimal).
    Property Get bg ' Return Type string
    End Property

    ' For a list view control, returns a DialogListColumns object that lets you query or modify the columns in Details mode.
    Property Get columns ' Return Type object:DialogListColumns 
    End Property

    ' Returns the number of items contained in the control (e.g. in a combo box, list box or list view, returns the number of items in the list).
    Property Get count ' Return Type int
    End Property

    ' Set or query the width of the control, in pixels.
    Property Get cx ' Return Type int
    End Property

    ' Set or query the height of the control, in pixels.
    Property Get cy ' Return Type int
    End Property

    ' Set or query the enabled state of the control. Returns True if the control is enabled, False if it's disabled. You can set this property to change the state.
    Property Get enabled ' Return Type bool
    End Property

    ' Set or query the color used for the text (foreground) of this control. This is in the format #RRGGBB (hexadecimal) or RRR,GGG,BBB (decimal).
    Property Get fg ' Return Type string
    End Property

    ' Set or query the input focus state of the control. Returns True if the control currently has input focus, False if it doesn't. Set to True to give the control input focus.
    Property Get focus ' Return Type bool
    End Property

    ' Set or query the control's label. Not all controls have labels - this will have no effect on controls (like the list view) that don't.
    Property Get label ' Return Type string orobject:Image 
    End Property

    ' For a list view control, lets you change or query the current view mode. Valid values are icon, details, smallicon, list.
    Property Get mode ' Return Type string
    End Property

    ' Set or query the read only state of an edit control.
    Property Get readonly ' Return Type bool
    End Property

    ' For a static text control set to ""image"" mode, you can set this property to rotate the displayed image. The value provided is the number of degrees from the image's initial orientation.
    Property Get rotate ' Return Type int
    End Property

    ' Set or query the font styles used to display this control's label. The string consists of zero or more characters; valid characters are b for bold and i for italics.
    Property Get style ' Return Type string
    End Property

    ' Set or query the color used for the text background (fill) of this control. This is in the format #RRGGBB (hexadecimal) or RRR,GGG,BBB (decimal).
    Property Get textbg ' Return Type string
    End Property

    ' Set or query the control's value. The meaning of this property depends on the type of the control:
    Property Get value ' Return Type string or bool or int or object:DialogListItem or object:Vector
    End Property

    ' Set or query the visible state of the control. Returns True if the control is visible and False if it's hidden. You can set this property to hide or show the control.
    Property Get visible ' Return Type bool
    End Property

    ' Set or query the left (x) position of the control, in pixels.
    Property Get x ' Return Type int
    End Property

    ' Set or query the top (y) position of the control, in pixels.
    Property Get y ' Return Type int 
    End Property

    ' Adds a new group to a list view control. Items you add to the list can optionally be placed in groups. Each group must have a unique ID.
    Function AddGroup(<string:name>,<int:id>,[<string:flags>]) ' Return Type int
    End Function

    ' Adds a new item to the control (list box, combo box or list view). The first parameter is the item's name, and the optional second parameter is a data value to associate with the item.
    Function AddItem (<string:name>,[<int:value>],[<int:groupid>] or <object:item>) ' Return Type int
    End Function

    ' This method is mainly for use with multiple-selection list box and list view controls. It lets you deselect individual items in the control while leaving other items selected (or unaffected).
    Function DeselectItem ' Return Type int
    End Function

    ' You can also specify -1 to deselect all items in the list box.
    Function DeselectItem ' Return Type int
    End Function

    ' Only applies to list view controls. By default group view is off; after adding groups with the AddGroup method, use EnableGroupView to turn group view on.
    Sub EnableGroupView
    End Sub

    ' Returns a DialogListGroup object representing the group with the specified ID that you've previous added to a list view control using the AddGroup method.
    Function GetGroupById ' Return Type object:DialogListGroup
    End Function

    ' Returns a DialogListItem object representing the item contained in the control at the specified index (list box, combo box or list view). Item 0 represents the first item in the list, item 1 the second, and so on.
    Function GetItemAt ' Return Type object:DialogListItem 
    End Function

    ' Returns a DialogListItem object representing the item contained in the control with the specified name (list box, combo box or list view). This method has two names (...Label and ...Name) for historical reasons, you can use either method name interchangeably).
    Function GetItemByLabel ' Return Type object:DialogListItem 
    End Function

    ' Inserts a new item in the control (list box, combo box or list view). The first parameter is the position to insert the item at (0 means the beginning of the list, 1 means the second position and so on). The second parameter is the item's name, and the optional third parameter is a data value to associate with the item.
    Function InsertItemAt(<int:position>,<string:name>,[<int:value>],[<int:groupid>] or <int:position>,<object:item>) ' Return Type int
    End Function

    ' Moves an existing item to a new location (list box, combo box or list view). The first parameter is the item to move (you can pass either its index or a DialogListItem object), and the second parameter is the new position the item should be moved to.
    Function MoveItem(<int:position> or <object:item>,<int:newposition>) ' Return Type  int
    End Function

    ' Removes the specified group from a list view control.
    Sub RemoveGroup
    End Sub

    ' Removes an item from the control (list box, combo box or list view). You can provide either the index of the item to remove (0 means the first item, 1 means the second and so on) or a DialogListItem object obtained from the GetItemAt or GetItemByName methods.
    Sub RemoveItem(<int:position> or <object:item>)
    End Sub

    ' Selects an item in the control. For a list box, combo box or list view, you can specify either the index of the item to select (0 means the first item, 1 means the second and so on) or a DialogListItem object obtained from the GetItemAt or GetItemByName methods.
    Function SelectItem(<int:position> or <object:item> or <string:tab>) ' Return Type int
    End Function

    ' Selects text within an edit control (or the edit field in a combo box control). The two parameters represent the start and end position of the desired selection. To select the entire contents, use 0 for the start and -1 for the end.
    Function SelectRange(<int:start> or <int:end> or <object:item1>,<object:item2>) ' Return Type object:Vector 
    End Function

    ' Sets the position of this control. The x and y coordinates are specified in pixels.
    Sub SetPos(<int:x>,<int:y>)
    End Sub

    ' Sets the position and size of the control, in a single operation. All coordinates are specified in pixels.
    Sub SetPosAndSize(<int:x>,<int:y>,<int:cx>,<int:cy>)
    End Sub

    ' Sets the size of this control. The cx (width) and cy (height) values are specified in pixels.
    Sub SetSize(<int:cx>,<int:cy>)
    End Sub
End Class

'The CustomFieldData object is provided to a rename script via the GetNewNameData.custom property. It provides access to the value of any custom fields that your script added to the Rename dialog.
Class CustomFieldData
End Class

'This object is provided to make it easier to deal with variables representing dates. It is obtained from the DOpusFactory.Date method as well as various properties in other objects.
Class Date ' Default Return Type date
    ' Get or set the day value of the date.
    Property Get day ' Return Type int
    End Property

    ' Get or set the hour value of the date.
    Property Get hour ' Return Type int
    End Property

    ' Get or set the minute value of the date.
    Property Get min  ' Return Type int
    End Property

    ' Get or set the month value of the date.
    Property Get month ' Return Type int
    End Property

    ' Get or set the milliseconds value of the date.
    Property Get ms ' Return Type int
    End Property

    ' Get or set the seconds value of the date.
    Property Get sec ' Return Type int
    End Property

    ' Get the day-of-the-week value of the date.
    Property Get wday  ' Return Type int
    End Property

    ' Get or set the year value of the date.
    Property Get year ' Return Type int
    End Property

    ' Adds the specified value to the date. The interpretation of the specified value is controlled by the type string:
    Sub Add(<int:value>,<string:type>)
    End Sub

    ' Returns a new Date object set to the same date as this one.
    Function Clone ' Return Type object:Date
    End Function
    ' Compares this date against the other date. The return value will be 0 (equal), 1 (greater) or -1 (less).
    Function Compare(<date:other>,[<string:type>],[<int:tolerance>]) ' Return Type int
    End Function

    ' Returns a formatted date or time string. The format and flags arguments are both optional.
    Function Format([<string:format>],[<string:flags>]) ' Return Type string
    End Function

    ' Returns a new Date object with the date converted from UTC (based on the local time zone).
    Function FromUTC ' Return Type object:Date
    End Function

    ' Resets the date to the current local date/time.
    Sub Reset
    End Sub

    ' Sets the value of this Date object to the supplied date.
    Sub Set
    End Sub

    ' Subtracts the specified value from the date. The parameters are the same as for the Add method.
    Sub Sub(<int:value>,<string:type>)
    End Sub

    ' Returns a new Date object with the date converted to UTC (based on the local time zone).
    Function ToUTC(none) ' Return Type object:Date
    End Function
End Class

'This object is used to display dialogs or popup menus. It is obtained from the Func.Dlg, Command.Dlg or DOpus.Dlg methods.
Class Dialog
    ' Specifies the buttons that are displayed at the bottom of the dialog. These buttons are used to close the dialog. The Show method returns a value indicating which button was chosen (and this value is also available in the result property).
    Property Get buttons ' Return Type string
    End Property

    ' This property uses either a Vector or an array of strings to provide a list of multiple options that can be shown to the user. The list can be presented in one of three ways:
    Property Get choices ' Return Type object:Vector(string) or array(string)
    End Property

    ' In a text entry dialog (i.e. the max property has been specified) setting confirm to True will require that the user types the entered text again (in a second text field) to confirm it (e.g. for a password).
    Property Get confirm ' Return Type bool
    End Property

    ' For script dialogs marked as resizable, this property lets you override the width of the dialog defined in the resource - although note you can't resize a dialog smaller than its initial size.
    Property Get cx ' Return Type int
    End Property

    ' For script dialogs marked as resizable, this property lets you override the height of the dialog defined in the resource - although note you can't resize a dialog smaller than its initial size.
    Property Get cy ' Return Type int
    End Property

    ' In a text entry dialog (i.e. the max property has been specified) this property allows you to initialize the text field with a default value.
    Property Get defvalue ' Return Type string
    End Property

    ' Allows you to change the default button (i.e. the action that will occur if the user hits enter) in the dialog. Normally the first button is the default - this has a defid of 1. The second button would have a defid of 2, and so on. If a dialog has more than one button then by definition the very last button is the ""cancel"" button, and this has a defid of 0.
    Property Get defid ' Return Type int
    End Property

    ' Set to True if you want a script dialog to run in “detached” mode, where your script provides its message loop.
    Property Get detach ' Return Type bool
    End Property

    ' Use this to cause the dialog to automatically disable another window when it's displayed. The user will be unable to click or type in the disabled window until the dialog is closed. Normally if you use this you would set this to the same value as the window property.
    Property Get disable_window ' Return Type object:Lister or object:Tab or object:Dialog or int
    End Property

    ' Displays one of several standard icons in the top-left corner of the dialog, which can be used, for example, to indicate the severity of an error condition. The valid values for this property are warning, error, info and question.
    Property Get icon ' Return Type string or object:Image 
    End Property

    ' In a text entry dialog, this property returns the text string that the user entered (i.e. once the Show method has returned).
    Property Get input ' Return Type string
    End Property

    ' Set this property to create a script dialog in a particular language (if one or more language overlays have been provided), rather than the currently selected language.
    Property Get language ' Return Type string
    End Property

    ' In conjunction with the choices property, this will cause the choices to be presented as a checkbox list. You can initialize this Vector or array with the same number of items as the choices property, and set each one to True or False to control the default state of each checkbox. Or, simply set this value to 0 to activate the checkbox list without having to initialize the state of each checkbox.
    Property Get list ' Return Type object:Vector(bool) or array(bool) or int
    End Property

    ' This property enables text entry in the dialog - a text field will be displayed allowing the user to enter a string. Set this property to the maximum length of the string you want the user to be able to enter (or 0 to have no limit).
    Property Get max ' Return Type int
    End Property

    ' In conjunction with the choices property, this will cause the choices to be presented as a popup menu rather than in a dialog. The menu will be displayed at the current mouse coordinates.
    Property Get menu ' Return Type object:Vector(int) or array(int) or int
    End Property

    ' Specifies the message text displayed in the dialog.
    Property Get message ' Return Type string
    End Property

    ' For script dialogs this property retrieves or sets the current dialog opacity level, from 0 (totally transparent) to 255 (totally opaque).
    Property Get opacity ' Return Type int
    End Property

    ' This is a collection of five options that will be displayed as checkboxes in the dialog. Unlike the choices / list scrolling checkbox list, these options are displayed as physical checkbox controls. By default the five checkboxes are uninitialized and won't be displayed, but if you assign a label to any of them they will be shown to the user.
    Property Get options ' Return Type collection:DialogOption 
    End Property

    ' In a text entry dialog, set this property to True to make the text entry field a password field. In a password field the characters the user enters are not displayed.
    Property Get password ' Return Type bool
    End Property

    ' When used with a script dialog this property lets you control the dialog's position on screen. Accepted values are:
    Property Get position ' Return Type string
    End Property

    ' By default, Opus checks the size and position of your dialog just before it opens and fixing them if they would place any of the dialog off-screen. Positioning a dialog off-screen is usually an accident caused by saving window positions on one system and restoring them on another with different monitor resolutions or arrangements. In the rare cases where you want your dialog to open off-screen, where the user cannot see some of all of it, set this property to False.
    Property Get position_fix ' Return Type bool
    End Property

    ' This property returns the index of the button chosen by the user to close the dialog. The left-most button is index 1, the next button is index 2, and so on. If a dialog has more than one button then by definition the last (right-most) button is the ""cancel"" button and so this will return index 0.
    Property Get result ' Return Type int
    End Property

    ' In a text entry dialog, set this property to True to automatically select the contents of the input field (as specified by the defvalue property) when the dialog opens.
    Property Get select ' Return Type bool
    End Property

    ' In a drop-down list dialog (one with the choices property set without either list or menu), this property returns the index of the item chosen from the drop-down list after the Show method returns.
    Property Get selection ' Return Type int
    End Property

    ' Set this property to True if the list of choices given by the choices property should be sorted alphabetically.
    Property Get sort ' Return Type bool
    End Property

    ' Lets you create a script dialog. The template property can be set to the name of the script dialog to display (as defined in your script resources), or a string that contains raw XML defining the dialog.
    Property Get template ' Return Type string
    End Property

    ' Specifies the title text of the dialog.
    Property Get title ' Return Type string
    End Property

    ' Set this property to True to make the dialog ""top level"", or False to allow it to go behind other non-top level windows.
    Property Get top ' Return Type bool
    End Property

    ' Set this property to True if you want the script dialog to generate close events in your message loop when the user clicks the window close button. You'll need to close the dialog yourself using the EndDlg method.
    Property Get want_close ' Return Type bool
    End Property

    ' Set this property to True if you want the script dialog to generate resize events in your message loop when the user resizes the dialog.
    Property Get want_resize ' Return Type bool
    End Property

    ' Use this to specify the parent window of the dialog. The dialog will appear centered over the top of the specified window. You can provide either a Lister or a Tab object, or another Dialog. If you are showing this dialog in response to the OnAboutScript event, you can also pass the value of the AboutData.window property.
    Property Get window ' Return Type object:Lister or object:Tab or object:Dialog or int
    End Property

    ' Specifies the x-position of a script dialog. Use the position property to control how the position is interpreted. After the dialog has been displayed you can change this property to move the dialog around on-screen.
    Property Get x ' Return Type int
    End Property

    ' Specifies the y-position of a script dialog. Use the position property to control how the position is interpreted. After the dialog has been displayed you can change this property to move the dialog around on-screen.
    Property Get y ' Return Type int
    End Property

    ' Creates a hotkey (or keyboard accelerator) for the specified key combination. When the user presses this key combination in your dialog, a hotkey event will be triggered.
    Sub AddHotkey(<string:name>,<string:key>)
    End Sub

    ' When creating a script dialog, calling this method creates the underlying dialog but does not display it. This lets you create the dialog and then initialize its controls before it is shown to the user.
    Sub Create
    End Sub

    ' Returns a Control object corresponding to one of the controls on a script dialog. The control is identified by its name, as defined in the script dialog resource.
    Function Control(<string:name>,[<string:dialog>],[<string:tab>]) ' Return Type object:Control 
    End Function

    ' Deletes a hotkey you previously created with the AddHotkey method.
    Sub DelHotkey
    End Sub

    ' Allows the user to drag and drop one or more files from your dialog (and drop them in another window or application).
    Function Drag(collection:Item,<string:actions>) ' Return Type string
    End Function

    ' Ends a script dialog running in detached mode. Normally dialogs end automatically when the user clicks the close button or another button that has its Close Dialog property set to True. This method lets you end a dialog under script control. The optional parameter specifies the result code that the Dialog.result property will return.
    Sub EndDlg
    End Sub
    ' Displays a ""Browse for Folder"" dialog letting the user select a folder. The optional parameters are:
    Function Folder(<string:title>,<string:default>,<bool:expand>,<object:window>) ' Return Type object:Path 
    End Function

    ' Returns a Msg object representing the most recent input event in a script dialog (only used in detached mode).
    Function GetMsg (<string:message>,<string:default> or <object:window>,<byref string:result>) ' Return Type object:Msg 
    End Function

    ' Displays a text entry dialog allowing the user to enter a string. The optional parameters are:
    Function GetString(<string:message>,<string:default>,<string:max>,<string:buttons>,<string:title>,<object:window>,<byref string:result>) ' Return Type string
    End Function

    ' Stops the specified timer. The timer must previously have been created by a call to the SetTimer method.
    Sub KillTimer
    End Sub

    ' Restores the previously saved position of a script dialog. The position must have previously been saved by a call to the SavePosition method.
    Sub LoadPosition(<string:id>,<string:type>)
    End Sub

    ' Displays a ""Browse to Open File"" dialog that lets the user select one or more files. The optional parameters are:
    Function Multi(<string:title>,<string:default>,<object:window>) ' Return Type collection:Item 
    End Function

    ' Displays a ""Browse to Open File"" dialog that lets the user select a single file. The optional parameters are:
    Function Open(<string:title>,<string:default>,<object:window>) ' Return Type object:Item 
    End Function

    ' Displays a dialog with one or more buttons. The optional parameters are:
    Function Request(<string:message>,<string:buttons>,<string:title>,<object:window>) ' Return Type int
    End Function

    ' Turns a previously detached dialog into a non-detached one, by taking over and running the default message loop. The RunDlg method won't return until the dialog has closed. You might use this if you created a dialog using Create, in order to initialize its controls, but don't actually want to run an interactive message loop.
    Function RunDlg(none) ' Return Type int
    End Function

    ' Saves the position (and size) of the dialog to your Opus configuration. The position can then be restored later on by a call to LoadPosition.
    Sub SavePosition
    End Sub

    ' Creates a timer that will generate a periodic timer event for your script. The period must be specified in milliseconds (e.g. 1000 would equal one second).
    Function SetTimer(<int:period>,<string:name>) ' Return Type string
    End Function

    ' Displays the dialog that has been pre-configured using the various properties of this object. See the properties section above for a full description of these.
    Function Show ' Return Type int
    End Function

    ' Displays a ""Browse to Save File"" dialog that lets the user select a single file or enter a new filename to save. The optional parameters are:
    Function Save(<string:title>,<string:default>,<object:window>,<string:type>) ' Return Type object:Path 
    End Function

    ' Used to change how custom dialogs are grouped with other Opus windows on the taskbar. Specify a group name to move the window into an alternative group, or omit the group argument to reset back to the default group. If one or more windows are moved into the same group, they will be grouped together, separate from other the default group.
    Function SetTaskbarGroup ' Return Type bool
    End Function

    ' Returns a Vars object that represents the variables that are scoped to this particular dialog. This allows scripts to use variables that persist from one use of the dialog to another.
    Function Vars ' Return Type object:Vars 
    End Function

    ' Allows a script dialog to monitor events in a folder tab. You will receive notifications of the requested events through your message loop.
    Function WatchTab(<object:Tab>,<string:events>,<string:id>) ' Return Type bool
    End Function
End Class

'The DialogListColumn object represents a column in a Details mode list view control in a script dialog. It's obtained by enumerating the DialogListColumns object.
Class DialogListColumn
    ' Returns or sets the column's name.
    Property Get name ' Return Type string
    End Property

    ' Set this property to True if you want this column to automatically resize when the list view is resized horizontally. Only one column can be set to auto-resize at a time.
    Property Get resize ' Return Type bool
    End Property

    ' Returns 1 if the list view is currently sorted forwards by this column, -1 if it's currently sorted backwards by this column, or 0 otherwise. Settings this property will re-sort the list.
    Property Get sort ' Return Type int
    End Property

    ' Returns or sets the column's width in pixels. Set it to -1 to automatically size the column to fit its content. You can automatically resize all columns at once using the DialogListColumns.AutoSize method.
    Property Get width ' Return Type int
    End Property
End Class

'The DialogListColumns object lets you query or modify the columns in a Details mode list view control in a script dialog. Use the Control.columns property to obtain a DialogListColumns object.
Class DialogListColumns
    ' Adds a new column to the list view, and returns the index of the new column.
    Function AddColumn(<string:name>) ' Return Type int
    End Function

    ' Automatically sizes all columns in the list view to fit their content.
    Sub AutoSize
    End Sub

    ' Deletes the specified column.
    Sub DeleteColumn(<int:index>)
    End Sub

    ' Returns a DialogListColumn object representing the column in the specified position.
    Function GetColumnAt(<int:index>) ' Return Type object:DialogListColumn
    End Function

    ' Inserts a new column in the list view at the specified position, and returns the index of the new column.
    Function InsertColumn(<string:name>,<int:position>) ' Return Type int
    End Function
End Class

'The DialogListGroup object represents a group in a list view control in a script dialog. It's returned by the Control.GetGroupById method.
Class DialogListGroup
    ' Returns or sets the expansion state of this group. The group must have been added as ""collapsible"" via the Control.AddGroup method.
    Property Get expanded ' Return Type bool
    End Property

    ' Returns the ID of this group.
    Property Get id ' Return Type int
    End Property

    ' Returns the name of this group.
    Property Get name ' Return Type string
    End Property
End Class

'The DialogListItem object represents an item in a combo box or list box control in a script dialog. It's returned by the Control.GetItemAt and Control.GetItemByName methods.
Class DialogListItem
    ' Set or query the color used for the background (fill) of this item. This is in the format #RRGGBB (hexadecimal) or RRR,GGG,BBB (decimal).
    Property Get bg ' Return Type string
    End Property

    ' For a list view control with checkboxes enabled, returns or sets the check state of the item.
    Property Get checked ' Return Type int
    End Property

    ' Returns or sets the optional data value associated with this item.
    Property Get data ' Return Type int
    End Property

    ' For a list view control, returns or sets the disable state of this item. When a list view item is disabled it appears ghosted and can't be selected or right-clicked.
    Property Get disabled ' Return Type bool
    End Property

    ' Set or query the color used for the text (foreground) of this control. This is in the format #RRGGBB (hexadecimal) or RRR,GGG,BBB (decimal).
    Property Get fg ' Return Type string
    End Property

    ' Returns or sets the list view group that this item is a member of.
    Property Get group ' Return Type int
    End Property

    ' For a list view control, returns or sets the icon associated with this item. You can specify the path of a file or folder to use its icon, or a file extension (e.g. "".txt"") to use a generic filetype icon. You can also set it to ""dir"", ""file"", ""ftp"" and ""ftps"" to use generic icons. You can also extract an icon from a DLL or EXE by providing the path of the file followed by a comma and then the icon index within the file.
    Property Get icon ' Return Type string
    End Property

    ' Returns the 0-based index of this item within the control.
    Property Get index ' Return Type index
    End Property

    ' Returns or sets the item's name.
    Property Get name ' Return Type string
    End Property

    ' Returns or sets the item's selection state. Mostly useful with multiple-selection list box controls.
    Property Get selected ' Return Type bool
    End Property

    ' Returns or sets the text style this item will be displayed in. You should provide a string containing one or more of the following flags: ""b"" (bold), ""i"" (italics), ""u"" (underline).
    Property Get style ' Return Type string
    End Property

    ' For a list view control in Details mode, returns a collection of strings that lets you query or change the text of the item's sub-items. There will be one string in the collection for each column in the list, excluding the first column.
    Property Get subitems ' Return Type collection:string
    End Property
End Class

'This object is used in conjunction with the Dialog object. It lets you specify a checkbox option that is added to the dialog.
Class DialogOption
    ' Set this to the desired label of the checkbox.
    Property Get label ' Return Type string
    End Property

    ' Set this to the desired initial state of the checkbox. When the Dialog.Show method returns, you can read this property to find out the state the user chose.
    Property Get state ' Return Type bool
    End Property
End Class

'This object represents a floating toolbar. The Toolbar object provides a collection that represents all instances of that toolbar that are currently floating.
Class Dock ' Default return int, This is a handle to the window of the floating toolbar. It is not particularly useful.
End Class

'This object provides metadata properties relating to document files. It is obtained from the Metadata object.
Class DocMeta
End Class

'This object is a helper object that you can use to create various other objects like Map and Vector. It is obtained from the DOpus.Create method.
Class DOpusFactory
    ' Returns a new Blob object, that lets you access and manipulate a chunk of binary data from a script. If no parameters are given the new Blob will be empty - you can set its size using the resize method - otherwise you can specify the initial size as a parameter.
    Function Blob(none or <int:size> or <byte, byte, ...> or <Blob:source>) ' Return Type object:Blob 
    End Function

    ' Creates a new BusyIndicator object, that lets you control the breadcrumbs bar busy indicator from your script.
    Function BusyIndicator ' Return Type object:BusyIndicator 
    End Function

    ' Creates a new Command object, that lets you run Opus commands from a script.
    Function Command ' Return Type object:Command
    End Function

    ' Creates a new Date object. If an existing Date object or date value is specified the new object will be initialized to that value, otherwise the date will be set to the current local time.
    Function Date(none or <variant:date>) ' Return Type object:Date 
    End Function

    ' Creates a new Map object. If no arguments are provided, the Map will be empty. Otherwise, the Map will be pre-initialized with the supplied key/value pairs. For example: Map(""firstname"",""fred"",""lastname"",""bloggs"");. The individual keys and values can be different types.
    Function Map(none or <variant:key>, <variant:value>...) ' Return Type object:Map 
    End Function

    ' Creates a new case-sensitive StringSet object. If no arguments are provided, the StringSet will be empty. Otherwise it will be pre-initialized with the supplied strings; for example: StringSet(""dog"",""cat"",""pony"");
    Function StringSet(none or <string>, ...) ' Return Type object:StringSet 
    End Function

    ' Creates a new case-insensitive StringSet object. If no arguments are provided, the StringSet will be empty. Otherwise it will be pre-initialized with the supplied strings.
    Function StringSetI(none or <string>, ...) ' Return Type object:StringSet 
    End Function

    ' Creates a new StringTools object, that provides helper functions for string encoding and decoding.
    Function StringTools ' Return Type object:StringTools 
    End Function

    ' Creates a new UnorderedSet object. If no arguments are provided the UnorderedSet will be empty. Otherwise it will be pre-initialized with the supplied elements.
    Function UnorderedSet(none or variants...) ' Return Type object:UnorderedSet 
    End Function

    ' Creates a new Vector object. If no arguments are provided, the Vector will be empty.
    Function Vector(none or <int:elements> or variants... or object:Vector or array) ' Return Type object:Vector 
    End Function
End Class

'The DPI object is a helper object that provides a number of methods and properties relating to the system DPI setting. It's returned via the DOpus.DPI property.
Class DPI
    ' Returns the system DPI setting as a “dpi value” (e.g. 96, 192).
    Property Get dpi ' Return Type int
    End Property

    ' Returns the DPI settings as a “scale factor” (e.g. 100, 125, 200).
    Property Get factor ' Return Type int
    End Property

    ' Divides the provided size by the system DPI; e.g. if the system DPI was set to 150%, DPI.Divide(60) would return 40.
    Function Divide ' Return Type int
    End Function

    ' Scales the provided size by the system DPI; e.g. if the system DPI was set to 200%, DPI.Scale(75) would return 150.
    Function Scale ' Return Type int
    End Function
End Class

'The Drive object provides information about a drive (hard drive, CD ROM, etc) on your system. A Vector of drives on your system can be obtained from the FSUtil.Drives method.
Class Drive ' Default Returns the root of the drive (e.g. C:\).
    ' Returns a FileSize object indicating the available free space on the drive.
    Property Get avail  ' Return Type object:FileSize 
    End Property

    ' Returns the bytes-per-cluster value for the drive.
    Property Get bpc  ' Return Type int
    End Property

    ' Returns a string representing the filesystem type.
    Property Get filesys  ' Return Type string
    End Property

    ' Returns a value representing filesystem flags for the drive.
    Property Get flags ' Return Type int
    End Property

    ' Returns a FileSize object indicating the total free space on the drive.
    Property Get free ' Return Type object:FileSize 
    End Property

    ' Returns the drive's label.
    Property Get label ' Return Type string
    End Property

    ' Returns a FileSize object indicating the total size of the drive.
    Property Get total ' Return Type object:FileSize 
    End Property

    ' Returns a string indicating the drive type (removable, fixed, remote, cdrom, ramdisk).
    Property Get type  ' Return Type string
    End Property
End Class

'This object provides metadata properties relating to executable (program) files. It is obtained from the Metadata object.
Class ExeMeta
End Class

'The Favorite object represents a favorite folder. It is retrieved by enumerating or indexing the Favorites object.
Class Favorite ' Default Return  name of the favorite as string
    ' Returns True if this is a sub-folder, False if it's a favorite folder or separator.
    Property Get folder ' Return Type bool
    End Property

    ' Returns True if this is a separator.
    Property Get separator ' Return Type bool
    End Property

    ' Returns the path this favorite folder refers to as a Path object.
    Property Get path ' Return Type object:Path 
    End Property

    ' Changes the name of this favorite folder. Note that changes you make to the list are not saved until you call the Favorites.Save method.
    Sub SetName(<string:name>)
    End Sub

    ' Changes the path this favorite folder refers to. Note that changes you make to the list are not saved until you call the Favorites.Save method.
    Sub SetPath(<string:path> or <object:Path>)
    End Sub
End Class

'The Favorites object holds a collection of all the defined favorite folders. It is retrieved from the DOpus.favorites method.
Class Favorites ' Default Return the Favorites object:Favorite
    ' Adds a new favorite folder to the favorites list. Note that changes you make to the list are not saved until you call the Save method.
    Function Add(<string:typeOrName>,<string:path>,<int:insertpos> or <object:Favorite>) ' Return Type object:Favorite or object:Favorites
    End Function

    ' Deletes the specified favorite or sub-folder. Note that changes you make to the list are not saved until you call the Save method.
    Sub Delete(<object:Favorite> or <object:Favorites>)
    End Sub

    ' Lets you locate a sub-folder one or more levels below the current one. The name parameter is the name or path and name of the sub-folder to look for (e.g. ""myfave"", ""pictures/local"", etc).
    Function Find(<string:name>,<int:index>) ' Return Type object:Favorites
    End Function

    ' Saves any changes you've made to the favorites list. Once you call this method changes you have made will be reflected in Preferences and the favorites list in Listers. Note that you can only call this method on the main ""root"" Favorites object obtained from the DOpus.favorites property
    Sub Save
    End Sub

    ' Changes the name of this sub-folder. Note that changes you make to the list are not saved until you call the Save method. You can only call this method on Favorites objects that refer to sub-folders, and not the main ""root"" folder.
    Sub SetName(<string:name>)
    End Sub
End Class

'This object lets you read or write binary data from or to a file. It is obtained from the FSUtil.OpenFile and Item.Open methods.
Class File' Default returns the full pathname of the file.
    ' Returns a Win32 error code that indicates the success or failure of the last operation. If the previous operation succeeded this will generally be 0.
    Property Get error ' Return Type int
    End Property

    ' Returns a FileSize object representing the size of this file, in bytes.
    Property Get size ' Return Type object:FileSize 
    End Property

    ' Returns a FileSize object representing the current position of the read or write cursor within this file, in bytes.
    Property Get tell ' Return Type object:FileSize
    End Property

    ' Closes the underlying file handle. After this call the File object is still valid but it can no longer read or write data.
    Sub Close
    End Sub

    ' Reads data from the file. If you provide a target Blob as the first parameter, the data will be stored in that Blob. Otherwise, a Blob will be created automatically.
    Function Read(<blob:target>,<int:size>) ' Return Type int or object:Blob 
    End Function

    ' Moves the read or write cursor within this file. The delta parameter specifies how many bytes to move - how this is interpreted depends on the optional method parameter:
    Function Seek(<int:delta>,<string:method>) ' Return Type object:FileSize 
    End Function

    ' Modifies the attributes of this file. You can either pass a string indicating the attributes to set, or a FileAttr object. When using a string, valid attributes are:
    Function SetAttr(object:FileAttr or <string:attributes>) ' Return Type bool
    End Function

    ' Modifies one or more of the file's timestamps. The create and access parameters are optional. If you wish to specify no change for a timestamp, specify 0.
    Function SetTime(<date:modify>,<date:create>,<date:access>) ' Return Type bool
    End Function

    ' Modifies one or more of the file's timestamps. The create and access parameters are optional. If you wish to specify no change for a timestamp, specify 0.
    Function SetTimeUTC(<date:modify>,<date:create>,<date:access>) ' Return Type bool
    End Function

    ' Truncates the file at the current position of the write cursor. You can use this in conjunction with the Seek method to pre-allocate a file's space on disk, for greater performance (i.e. seek to the final size of the file, truncate at that point, and then seek back to the start and write the data).
    Function Truncate ' Return Type bool
    End Function

    ' Writes data from the specified Blob (or array) to the file. By default the entire contents of the Blob will be written - you can use the optional from parameter to specify the source byte offset, and the size parameter to specify the number of bytes to write.
    Function Write(<blob:source>,<int:from>,<int:size>)' Return Type int
    End Function
End Class

'This object represents file attributes (like read only, archived, etc). It used by the Item and Format objects, and can be created by the FSUtil.NewFileAttr method.
Class FileAttr ' Default Returns a string representing the attributes that are set (similar to the format displayed in the Attr column in the file display).
    ' A file or directory that has changes which need archiving. The A bit is usually set on new or modifies files, and may then be cleared by backup software after it has added the changes to a backup.
    Property Get a ' Return Type bool
    End Property

    '
    Property Get archive ' Return Type bool
    End Property

    ' A file or directory that is compressed. For a file, all of the data in the file is compressed. For a directory, compression is the default for newly created files and subdirectories.
    Property Get c ' Return Type bool
    End Property

    '
    Property Get compressed ' Return Type bool
    End Property

    ' A file or directory that is encrypted. For a file, all data streams in the file are encrypted. For a directory, encryption is the default for newly created files and subdirectories.
    Property Get e ' Return Type bool
    End Property

    '
    Property Get encrypted ' Return Type bool
    End Property

    ' The file or directory is hidden. It is not included in an ordinary directory listing.
    Property Get h ' Return Type bool
    End Property

    '
    Property Get hidden ' Return Type bool
    End Property

    ' The file or directory is not to be indexed by the content indexing service.
    Property Get i ' Return Type bool
    End Property

    '
    Property Get nonindexed ' Return Type bool
    End Property

    ' The data of a file is not available immediately. This attribute indicates that the file data is physically moved to offline storage. This attribute is used by Remote Storage, which is the hierarchical storage management software. Applications should not arbitrarily change this attribute.
    Property Get o ' Return Type bool
    End Property

    '
    Property Get offline ' Return Type bool
    End Property

    ' The data of the file is to be kept available at all times; it should not be offloaded to offline storage.
    Property Get p ' Return Type bool
    End Property

    '
    Property Get pinned ' Return Type bool
    End Property

    ' A file that is read-only. Applications can read the file, but cannot write to it or delete it. This attribute is not honored on directories.
    Property Get r ' Return Type bool
    End Property

    '
    Property Get readonly ' Return Type bool
    End Property

    ' A file or directory that the operating system uses a part of, or uses exclusively.
    Property Get s ' Return Type bool
    End Property

    '
    Property Get system ' Return Type bool
    End Property

    ' Assigns a new set of attributes to this object. You can pass another FileAttr object, or a string (e.g. ""hsr"").
    Sub Assign(object:FileAttr or string)
    End Sub

    ' Given a single character representing an attribute (e.g. ""a"") this method returns the name of the attribute in the user's current language (e.g. ""Archive"").
    Function AttrName(string) ' Return Type string
    End Function

    ' Clears (turns off) the specified attributes in this object. You can pass another FileAttr object, or a string representing the attributes to turn off.
    Sub Clear(object:FileAttr or string)
    End Sub

    ' Sets (turns on) the specified attributes in this object. You can pass another FileAttr object, or a string representing the attributes to turn on.
    Sub Set(object:FileAttr or string)
    End Sub

    ' Returns a string representing the attributes that are set (similar to the format displayed in the Attr column in the file display).
    Function ToString ' Return Type string
    End Function
End Class

'This object exposes information about a file group (when a Tab is set to group by a particular column). It is used by the Item and Tab objects.
Class FileGroup ' Default Returns the name of the group as string.
    ' Returns True if the group is currently collapsed.
    Property Get collapsed ' Return Type bool
    End Property

    ' Returns the number of items in this group. Note that groups can be empty; empty groups are not displayed in the file display but will still be returned by the Tab.filegroups property.
    Property Get count ' Return Type int
    End Property

    ' Returns the id number of this group. Id numbers are arbitrary - you shouldn't place any meaning on the actual value, but you can compare the id fields as an easy way to tell if two items are in the same group.
    Property Get id ' Return Type int
    End Property

    ' Returns a collection of Item objects that represents all the files and folders in this group.
    Property Get members ' Return Type collection:Item 
    End Property

    ' Returns a string indicating the collation type of the group.
    Property Get type ' Return Type string
    End Property
End Class

'This object is used to represent a size in bytes (mainly because ActiveX scripting doesn't have proper support for 64 bit integers). It is used by the Item and TabStats objects.
Class FileSize ' Default Returns the number of bytes represented by this FileSize object as a string.
    ' Returns the number of bytes as a currency value. This is a 64 bit data type but it is stored as a fractional value, so you must multiply the returned value by 10000 to obtain the actual byte size.
    Property Get cy ' Return Type currency
    End Property

    ' Returns the number of bytes as an automatically formatted string (e.g. if the FileSize value is 1024, the string 1 KB would be returned).
    Property Get fmt ' Return Type string
    End Property

    ' Returns the highest (most significant) 32 bits of the file size.
    Property Get high  ' Return Type int
    End Property

    ' Returns the lowest (least significant) 32 bits of the file size.
    Property Get low ' Return Type int
    End Property

    ' Adds the supplied value to the value of this FileSize object. You can pass a string, int or currency type, or another FileSize object.
    Sub Add(variant)
    End Sub

    ' Clones this FileSize object and returns a new one set to the same value.
    Function Clone ' Return Type object:FileSize
    End Function

    ' Compares the supplied value with the value of this FileSize object. The return value will be 0 (equal), 1 (greater) or -1 (less).
    Function Compare(variant) ' Return Type int
    End Function

    ' Divides the value of this FileSize object with the supplied value. You can pass a string, int or currency type, or another FileSize object.
    Sub Div(variant)
    End Sub

    ' Multiplies the value of this FileSize object with the supplied value. You can pass a string, int or currency type, or another FileSize object.
    Sub Mult(variant)
    End Sub

    ' Sets the FileSize to the supplied value. You can pass a string, int or currency type, or another FileSize object. You can also pass a Blob consisting of exactly 1, 2, 4 or 8 bytes, in which case the data contained in the Blob will be used to form the number.
    Sub Set(variant)
    End Sub

    ' Subtracts the supplied value from the value of this FileSize object. You can pass a string, int or currency type, or another FileSize object. Note that the FileSize object is unsigned and so the value cannot go below zero.
    Sub Sub(variant)
    End Sub

    ' Returns a Blob containing the bytes that make up the current value. By default 8 bytes will be copied to the Blob (the full 64 bit number) but you can pass an alternative number of bytes (1, 2 or 4) as a parameter to truncate the value.
    Function ToBlob(int) ' Return Type object:Blob 
    End Function
End Class

'This object represents a file type group (as configured in the File Type Groups section of the file type editor).
Class FiletypeGroup ' Default Returns the internal name of this group as string.
    ' Returns the display name of this group.
    Property Get display_name ' Return Type string
    End Property

    ' Returns the tiles mode definition string for this group.
    Property Get tiles ' Return Type string
    End Property

    ' Returns the tooltip definition string for this group.
    Property Get tooltip ' Return Type string
    End Property

    ' Tests the filename (or extension) for membership of this group. Returns True if the file is a member of the group, or False if it is not.
    Function MatchExt(<string:filename>) ' Return Type bool
    End Function
End Class

'This object represents a collection of one or more file type groups.
Class FiletypeGroups 'Default Lets you enumerate the file type groups as FiletypeGroup object.
    ' Searches the file type group collection for the named group.
    Function GetGroup(<string:group>) ' Return Type object:FiletypeGroup or bool (False)
    End Function

    ' Returns a new FiletypeGroups object containing the subset of groups that the specified filename (or file extension) is a member of. You would normally only call this method on the object returned by the DOpus.filetypegroups property.
    Function MatchExt(<string:filename>) ' Return Type object:FiletypeGroups
    End Function

    ' Returns the translated name of the named built-in file type group.
    Function Translate(<string:group>) ' Return Type string 
    End Function
End Class

'This object lets a script enumerate the contents of a folder. It is obtained using the FSUtil.ReadDir method.
Class FolderEnum
    ' True if the enumeration is complete, otherwise False.
    Property Get complete ' Return Type bool
    End Property

    ' If an error occurs this will return the error code. It will return 0 on success.
    Property Get error ' Return Type int
    End Property

    ' Closes the underlying file system handle used to perform the enumeration. You might call this method if you want to delete the folder you just enumerated. After this method is called the complete property will return True.
    Sub Close
    End Sub

    ' Returns the next item in the enumeration.
    Function Next(<int:count> or <Vector:vector>) ' Return Type object:Item or object:Vector
    End Function
End Class

'This object provides metadata properties relating to font files. It is obtained from the Metadata object.
Class FontMeta
    ' The character set.
    Property Get charset ' Return Type int
    End Property

    ' The clipping precision.
    Property Get clipprecision ' Return Type int
    End Property

    ' The angle, in tenths of degrees, between the escapement vector and the x-axis of the device.
    Property Get escapement ' Return Type int
    End Property

    ' The typeface name of the font.
    Property Get fontname ' Return Type string
    End Property

    ' The height, in logical units, of the font's character cell or character.
    Property Get height ' Return Type int
    End Property

    ' An italic font if set to True.
    Property Get italic ' Return Type bool
    End Property

    ' The angle, in tenths of degrees, between each character's base line and the x-axis of the device.
    Property Get orientation ' Return Type int
    End Property

    ' The output precision.
    Property Get outprecision ' Return Type int
    End Property

    ' The pitch and family of the font.
    Property Get pitchandfamily ' Return Type int
    End Property

    ' The output quality.
    Property Get quality ' Return Type int
    End Property

    ' A strikeout font if set to True.
    Property Get strikeout ' Return Type bool 
    End Property

    ' An underlined font if set to True.
    Property Get underline ' Return Type bool
    End Property

    ' The weight of the font in the range 0 through 1000.
    Property Get weight ' Return Type int
    End Property

    ' The average width, in logical units, of characters in the font.
    Property Get width ' Return Type int
    End Property
End Class

'This object provides information about the display format in a tab. It is obtained from the Tab.format property.
Class Format
    ' Returns True if folders are always sorted alphabetically, False if otherwise.
    Property Get alpha_folders ' Return Type bool
    End Property

    ' Returns True if column width auto-sizing is enabled, False if otherwise.
    Property Get autosize ' Return Type bool
    End Property

    ' Returns a collection of Column objects that represent all the individual columns currently added to the display.
    Property Get columns ' Return Type collection:Column 
    End Property

    ' Returns a Vector of strings representing the explanation of the current folder format (the same text visible when hovering the mouse over the format lock icon in the status bar).
    Property Get format_explain ' Return Type Vector:string
    End Property

    ' Returns a string that indicates the state of the option to automatically calculate folder sizes. The string returned will be one of default, on or off.
    Property Get getsizes ' Return Type string
    End Property

    ' If grouping is enabled, returns the name of the column that the list is grouped by.
    Property Get group_by ' Return Type string
    End Property

    ' Returns True if the Individual groups option is enabled.
    Property Get group_individual ' Return Type bool
    End Property

    ' Returns True if the groups are sorted in reverse order.
    Property Get group_reverse ' Return Type bool
    End Property

    ' Returns a FileAttr object indicating the file attributes that are hidden (any items with these attributes set will be hidden from the display).
    Property Get hide_attr ' Return Type object:FileAttr 
    End Property

    ' Returns the wildcard pattern of folders that are hidden from the display.
    Property Get hide_dirs ' Return Type string
    End Property

    ' Returns True if the current hide_dirs pattern is using regular expressions.
    Property Get hide_dirs_regex ' Return Type bool
    End Property

    ' Returns True if filename extensions are hidden, or False if they are displayed.
    Property Get hide_ext ' Return Type bool
    End Property

    ' Returns the wildcard pattern of files that are hidden from the display.
    Property Get hide_files ' Return Type string
    End Property

    ' Returns True if the current hide_files pattern is using regular expressions.
    Property Get hide_files_regex ' Return Type bool
    End Property

    ' Returns a FileAttr object indicating the folder attributes that are hidden (any folders with these attributes set will be hidden from the display). If the separate folder attribute filter is disabled this property will return the string ""off"".
    Property Get hide_folder_attr ' Return Type object:FileAttr or string
    End Property

    ' Returns the filename prefixes that are ignored when sorting the list.
    Property Get ignore_prefix ' Return Type string
    End Property

    ' Returns True if the folder format is locked in the tab.
    Property Get locked ' Return Type bool
    End Property

    ' Returns True if manual sorting is enabled.
    Property Get manual_sort ' Return Type bool
    End Property

    ' If manual sorting is active, returns the name of the current sort order (if it has one).
    Property Get manual_sort_name ' Return Type string
    End Property

    ' If manual sort is active, returns a SortOrder object which lets you query and change the sort order.
    Property Get manual_sort_order ' Return Type object:SortOrder 
    End Property

    ' Returns a string indicating the current file/folder mixing type. The string returned will be one of mixed, files (files first) or dirs (folders first).
    Property Get mix_type ' Return Type string
    End Property

    ' Returns True if filenames and extensions are sorted separately.
    Property Get name_ext ' Return Type bool
    End Property

    ' Returns True if numeric name sorting is enabled.
    Property Get numeric_name ' Return Type bool
    End Property

    ' Returns True if the over-all sort order is reversed.
    Property Get reverse_sort ' Return Type bool
    End Property

    ' Returns a FileAttr object indicating the file attributes that are shown (only items with these attributes set will be shown in the display).
    Property Get show_attr ' Return Type object:FileAttr
    End Property

    ' Returns the wildcard pattern of folders that are shown (only folders matching this pattern will be shown).
    Property Get show_dirs ' Return Type string
    End Property

    ' Returns True if the current show_dirs pattern is using regular expressions.
    Property Get show_dirs_regex ' Return Type bool
    End Property

    ' Returns the wildcard pattern of files that are shown.
    Property Get show_files ' Return Type string
    End Property

    ' Returns True if the current show_files pattern is using regular expressions.
    Property Get show_files_regex ' Return Type bool
    End Property

    ' Returns a FileAttr object indicating the folder attributes that are shown (only folders with these attributes set will be shown in the display). If the separate folder attribute filter is disabled this property will return the string ""off"".
    Property Get show_folder_attr ' Return Type object:FileAttr or string
    End Property

    ' Returns True if the name column is sorted by filename extension rather than filename.
    Property Get sort_ext ' Return Type bool
    End Property

    ' Returns a Column object representing the current sort field.
    Property Get sort_field ' Return Type object:Column 
    End Property

    ' Returns the current view mode as a string. The returned string will be one of large_icons, small_icons, list, details, power, thumbnails or tile.
    Property Get view ' Return Type string
    End Property

    ' Returns True if word sorting is enabled.
    Property Get word_sort ' Return Type bool
    End Property

    ' The first time a script accesses a particular Format object, a snapshot is taken of the tab's format. If the script then makes changes to that tab (e.g. it changes the sort field, etc), these changes will not be reflected by the object. To re-synchronize the object with the tab, call the Format.Update method.
    Sub Update
    End Sub
End Class

'This object provides various utility methods relating to file system activity. It is obtained from the DOpus.FSUtil property.
Class FSUtil
    ' Compares the two provided path strings for equality - returns True if the two paths are equal, or False if otherwise.
    Function ComparePath(<string:path1>,<string:path2>,<string:flags>) ' Return Type bool
    End Function

    ' Retrieves the display name of a path. This is the form of a path that is intended to be displayed to the user, rather than used internally by Opus. For example, for a library path it will strip off the internal ?xxxxxxx notation that Opus uses to identify library member folders.
    Function DisplayName(<string:path>,<string:flags>) ' Return Type string
    End Function

    ' Returns a Vector of Drive objects, one for each drive on the system.
    Function Drives ' Return Type Vector:Drive 
    End Function

    ' Returns True if the specified file, folder or device exists, or False otherwise.
    Function Exists(<string:path>) ' Return Type bool
    End Function

    ' Returns the localized text description for a system error code.
    Function GetErrorMsg(<int:error>) ' Return Type string
    End Function

    ' Creates an Item object for the specified file path.
    Function GetItem(<string:path>) ' Return Type object:Item 
    End Function

    ' Returns a Metadata object representing the metadata for the specified file.
    Function GetMetadata(<string:path>) ' Return Type object:Metadata 
    End Function

    ' Returns the value of one or more shell properties for the specified file. The file path must be provided as the first parameter. If the second parameter is the name (or PKEY) of the property to retrieve, the value of the property will be returned as the return value from this method.
    Function GetShellProperty(<string:path>,<string:property> or <Map:properties>,<string:type>) ' Return Type variant
    End Function

    ' Returns a Vector of ShellProperty objects which represents all the possible shell properties available on the system. You can optionally provide a wildcard pattern as the first argument - if you do, only properties whose names match the supplied pattern will be returned.
    Function GetShellPropertyList(<string:pattern>,<string:type>) ' Return Type object:ShellProperty
    End Function

    ' Returns a string indicating the type of the specified file path. The string will be either file, dir or invalid if the path doesn't exist. The optional flags argument is used to control the behavior with archives - normally an archive will be reported as dir, but if you specify ""a"" for the flags parameter it will be reported as file.
    Function GetType(<string:path>,<string:flags>) ' Return Type string
    End Function

    ' Calculates a checksum for the specified file or Blob. By default the MD5 checksum is calculated, but you can use the optional type parameter to change the checksum algorithm - valid values are ""md5"", ""sha1"", ""sha256"", ""sha512"", ""crc32"", ""crc32_php"", and ""crc32_php_rev"".
    Function Hash(<string:path> or <object:Blob>,<string:type>) ' Return Type string or object:Vector 
    End Function

    ' Creates a new FileAttr object, which represents file attributes. You can initialize the new object by passing either a string representing the attributes to turn on (e.g. ""hsr"") or another FileAttr object. If you don't pass a value the new object will default to all attributes turned off.
    Function NewFileAttr(<attributes>) ' Return Type object:FileAttr 
    End Function

    ' Creates a new FileSize object, which makes it easier to handle 64 bit file sizes. You can initialize this with a number of data types (int, string, currency, another FileSize object, or a Blob containing exactly 1, 2, 4 or 8 bytes).
    Function NewFileSize(<string:""s"">,<size>) ' Return Type object:FileSize 
    End Function

    ' Creates a new Path object initialised to the provided path string.
    Function NewPath(<string:path>) ' Return Type object:Path
    End Function

    ' Creates a new Wild object. If a pattern and flags are provided the pattern will be parsed automatically, otherwise you must call the Parse method to parse a pattern before using the object.
    Function NewWild(<string:pattern>,<string:flags>) ' Return Type object:Wild 
    End Function

    ' Opens or creates a file and returns a File object that lets you access its contents as binary data. You can pass a string or Path object representing a file to open, and you can also pass an existing Blob object to create a File object that gives you read/write stream access to a chunk of memory.
    Function OpenFile(<string:path> or <object:Blob>,<string:mode>,<object:window> or <string:elevation>) ' Return Type object:File 
    End Function

    ' Returns a FolderEnum object that lets you enumerate the contents of the specified folder.
    Function ReadDir(<string:path>,<string:flags>) ' Return Type object:FolderEnum 
    End Function

    ' Resolves the specified library or file collection path to its underlying file system path. This method can also be used to resolve a folder alias, as well as application paths in the form {apppath|appname}.
    Function Resolve(<string:path>,<string:flags>) ' Return Type object:Path
    End Function

    ' Returns True if this path refers to the same drive or partition as the supplied path. The optional flags are:
    Function SameDrive(<string:path>,<string:flags>) ' Return Type bool
    End Function
End Class

'This object is passed to a script function (via ClickData.func) or script-defined internal command (via ScriptCommandData.func). It provides information relating to the function invocation (source and destination tabs, arguments, etc).
Class Func
    ' Returns an Args object that provides access to any arguments given on the command line that invoked this script. This is used when the script has added an internal command to Opus. A command line template can be provided when the command is added, and any arguments the user provides on the command line for the script command will be available via this object.
    Property Get args ' Return Type object:Args 
    End Property

    ' Returns a Map object that provides keyword lookup for each of the arguments given on the command line. An argument will only be present in the Map if it was used on the command line, so you can easily check which arguments are present using the Map.exists() method.
    Property Get argsmap ' Return Type object:Map 
    End Property

    ' This property returns a pre-filled Command object that can be used to run commands against the source and destination tabs. Using this object is the equivalent of calling DOpusFactory.Command and setting the source and destination tabs manually.
    Property Get command ' Return Type object:Command 
    End Property

    ' This object represents the default destination tab for the function.
    Property Get desttab ' Return Type object:Tab 
    End Property

    ' Returns True if the command was invoked via a drag-and-drop operation.
    Property Get fromdrop ' Return Type bool
    End Property

    ' Returns True if the command was invoked via the keyboard (i.e. via a hotkey rather than a button).
    Property Get fromkey ' Return Type bool
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the command was invoked.
    Property Get qualifiers ' Return Type string
    End Property

    ' This object represents the default source tab for the function.
    Property Get sourcetab ' Return Type object:Tab
    End Property

    ' If this button was run from the standalone image viewer, this object represents the viewer window.
    Property Get viewer ' Return Type object:Viewer 
    End Property

    ' Creates a new Dialog object, that lets you display dialogs and popup menus. The dialog's window property will be automatically assigned to the source tab.
    Function Dlg ' Return Type object:Dialog
    End Function
End Class

'This object lets you access information about the global filter settings (configured on the Folders / Global Filters page in Preferences).
Class GlobalFilters
    ' Returns True if the global wildcard filters are enabled.
    Property Get enable ' Return Type bool
    End Property

    ' Returns the global filename filter wildcard pattern. If the wildcard is configured to use regular expressions, it will have a regex: prefix in front of the pattern.
    Property Get file ' Return Type string
    End Property

    ' Returns the global folder filter wildcard pattern. If the wildcard is configured to use regular expressions, it will have a regex: prefix in front of the pattern.
    Property Get folder ' Return Type string
    End Property

    ' Returns True if the global option to hide hidden files is on.
    Property Get hidehidden ' Return Type bool
    End Property

    ' Returns True if the global option to hide operating system files is on.
    Property Get hidesystem ' Return Type bool
    End Property
End Class

'This object represents an image file or icon loaded from disk that can be displayed in a script dialog.
Class Image
    ' Returns the bit count of the loaded image.
    Property Get bitcount ' Return Type int
    End Property

    ' Returns the height of the loaded image.
    Property Get height ' Return Type int
    End Property

    ' Returns the width of the loaded image.
    Property Get width ' Return Type int
    End Property
End Class

'This object provides metadata properties relating to image files. It is obtained from the Metadata object.
Class ImageMeta
End Class

'This object represents a file or a folder. It can be returned from various methods of the Tab object, when enumerating a folder using the FSUtil.ReadDir method, and is used to provide files for a command to act on using the Command object.
Class Item ' Default Returns the full pathname of the item (i.e. path plus filename).

    ' Returns the ""last accessed"" date, in local time.
    Property Get access ' Return Type date
    End Property

    ' Returns the ""last accessed"" date, in UTC.
    Property Get access_utc ' Return Type date
    End Property

    ' Returns the item attributes. This value is a series of flags that are logically OR'd together. The attributes supported by Opus are:
    Property Get attr ' Return Type int
    End Property

    ' Returns the item attributes as a string, as displayed in the file display.
    Property Get attr_text ' Return Type string
    End Property

    ' Returns True if the item was checked (in checkbox mode), or False otherwise.
    Property Get checked ' Return Type bool
    End Property

    ' Returns the ""creation"" date, in local time.
    Property Get create ' Return Type date
    End Property

    ' Returns the ""creation"" date, in UTC.
    Property Get create_utc ' Return Type date
    End Property

    ' For Item objects obtained from a Viewer, this property is True if the item represents the currently displayed image and False otherwise.
    Property Get current ' Return Type bool
    End Property

    ' Returns the display name of the item. Only a few items have a display name that is different to their actual name - some examples are certain system folders (like C:\Users which might have a translated display name in non-English locales).
    Property Get display_name ' Return Type string
    End Property

    ' Returns the filename extension.
    Property Get ext ' Return Type string
    End Property

    ' Returns the filename extension, taking multi-part extensions into account. For example, a file called ""file.part1.rar"" might return "".rar"" for ext but "".part1.rar"" for ext_m.
    Property Get ext_m ' Return Type string
    End Property

    ' Returns True if the item failed when used by a command. This is only meaningful in conjunction with the Command.files collection - once the command has returned, this property will indicate success or failure on a per-file basis.
    Property Get failed ' Return Type bool
    End Property

    ' Returns a FileAttr object that represents the item's attributes.
    Property Get fileattr ' Return Type object:FileAttr
    End Property

    ' If the file display this item came from is grouped by a particular column, this property returns a FileGroup object representing the group the item is in. If the item has no group this will return an empty string.
    Property Get filegroup ' Return Type object:FileGroup 
    End Property

    ' For Item objects obtained from a file display, this property is True if the object represents the item with focus, and False otherwise. Only one item can have focus at a time. The item with focus is typically shown with an outline around it, and is usually the last item which was clicked on, or which was moved to with the keyboard. The item with focus is often also one of the selected items, but not always; selection and focus are two separate things.
    Property Get focus ' Return Type bool
    End Property

    ' Returns True for folder items if their size has been calculated by, for example, the GetSizes command. If False, the size property will be unreliable for folders.
    Property Get got_size ' Return Type bool
    End Property

    ' Returns a Vector of FiletypeGroup objects representing any and all file type groups that this file is a member of.
    Property Get groups ' Return Type Vector:FiletypeGroup 
    End Property

    ' Similar to the groups property, except a FiletypeGroups object is returned instead of a Vector.
    Property Get groupsobject ' Return Type object:FiletypeGroups 
    End Property

    ' This is a unique ID for the item; it is used internally by Opus.
    Property Get id ' Return Type int
    End Property

    ' Returns True if the item represents a folder, and False for a file.
    Property Get is_dir ' Return Type bool
    End Property

    ' Returns True if the item is a junction to another folder.
    Property Get is_junction ' Return Type bool
    End Property

    ' Returns True if the item is a reparse point.
    Property Get is_reparse ' Return Type bool
    End Property

    ' Returns True if the item is a symbolic link.
    Property Get is_symlink ' Return Type bool
    End Property

    ' Returns a Metadata object that provides access to the item's metadata.
    Property Get metadata ' Return Type object:Metadata 
    End Property

    ' Returns the ""last modified"" date, in local time.
    Property Get modify ' Return Type date
    End Property

    ' Returns the ""last modified"" date, in UTC.
    Property Get modify_utc ' Return Type date
    End Property

    ' Returns the name of the item.
    Property Get name ' Return Type string
    End Property

    ' Returns the filename ""stem"" of the item. This is the name of the item with the filename extension removed. It will be the same as the name for folders.
    Property Get name_stem ' Return Type string
    End Property

    ' Returns the filename ""stem"" of the item, taking multi-part extensions into account. For example, a file called ""file.part1.rar"" might return ""file.part1"" for name_stem but ""file"" for name_stem_m.
    Property Get name_stem_m ' Return Type string
    End Property

    ' Returns the path of the item's parent folder. This does not include the name of the item itself, which can be obtained via the name property.
    Property Get path ' Return Type object:Path
    End Property

    ' Returns the ""real"" path of the item. For items located in virtual folders like Libraries or Collections, this lets you access the item's underlying path in the real file system. The realpath property includes the full path to the item, including its own name.
    Property Get realpath ' Return Type object:Path
    End Property

    ' Returns True if the item was selected, or False otherwise.
    Property Get selected ' Return Type bool
    End Property

    ' Returns the short path of the item, if it has one. Note that short paths are disabled by default in Windows 10.
    Property Get shortpath ' Return Type object:Path
    End Property

    ' Returns the size of the item as a FileSize object.
    Property Get size ' Return Type object:FileSize 
    End Property

    ' Tests the file for membership of the specified file type group.
    Function InGroup(<string:group>) ' Return Type bool
    End Function

    ' This method returns a Vector of strings representing any labels that have been assigned to the item.
    Function Labels(<string:category>,<string:flags>) ' Return Type Vector:string
    End Function

    ' Opens this file and returns a File object that lets you access its contents as binary data.
    Function Open(<string:mode>,<object:window>) ' Return Type object:File 
    End Function

    ' Returns the value of the specified shell property for the item. The property argument can be the property's PKEY or its name.
    Function ShellProp(<string:property>,<string:type>) ' Return Type variant
    End Function

    ' Updates the Item object from the file on disk. You might use this if you had run a command to change an item's timestamp or attributes, and wanted to retrieve the new information.
    Sub Update
    End Sub
End Class

'This object represents a Lister window.
Class Lister
    ' Returns a Tab object representing the currently active (source) tab.
    Property Get activetab ' Return Type object:Tab 
    End Property

    ' Lister window bottom-edge coordinate.
    Property Get bottom ' Return Type int
    End Property

    ' Returns the custom title of the Lister (if any) as set by the Set LISTERTITLE command. This may be an empty string. The title property returns the actual window title.
    Property Get custom_title ' Return Type string
    End Property

    ' Returns a Tab object representing the current destination tab (in a dual-display Lister).
    Property Get desttab ' Return Type object:Tab 
    End Property

    ' Indicates whether the Lister is in dual-display mode or not. Possible values are:
    Property Get dual ' Return Type int
    End Property

    ' Returns the current split percentage of the dual displays (e.g. 50 indicates they are evenly sized).
    Property Get dualsize ' Return Type int
    End Property

    ' Returns True if this Lister is currently the foreground (active) window.
    Property Get foreground ' Return Type bool
    End Property

    ' Returns True if this Lister is currently the active Lister (foreground window), or was the most recently active Lister.
    Property Get lastactive ' Return Type bool
    End Property

    ' Provides the name of the Lister layout that this Lister came from (if any).
    Property Get layout ' Return Type string
    End Property

    ' Lister window left-edge coordinate.
    Property Get left ' Return Type int
    End Property

    ' Indicates whether the metadata pane is currently open or not. Possible values are:
    Property Get metapane ' Return Type int
    End Property

    ' Lister window right-edge coordinate.
    Property Get right ' Return Type int
    End Property

    ' Returns the state of a single-display mode Lister:
    Property Get state ' Return Type string
    End Property

    ' Returns the name of the Lister style which was last applied to the Lister, or an empty string if there is none. This is just the last style which was loaded and does not mean the Lister still looks the same; the user may have opened or closed panels and made other changes via other methods in the time since the style was applied.
    Property Get style ' Return Type string
    End Property

    ' Returns a collection of Tab objects that represent all tabs in this Lister. In a dual-display Lister this includes tabs in both the left and right file displays.
    Property Get tabs ' Return Type collection:Tab
    End Property

    ' Returns the name of the Folder Tab Group which was last loaded into the left half of the Lister, or an empty string if no group has been loaded.
    Property Get tabgroupleft ' Return Type string
    End Property

    ' Similar to tabgroupleft, above, but for the right half of the Lister (if any).
    Property Get tabgroupright ' Return Type string
    End Property

    ' Returns a collection of Tab objects that represent the tabs in the left/top side of a dual-display Lister. In a single-display Lister this is equivalent to all the tabs in the Lister.
    Property Get tabsleft ' Return Type collection:Tab
    End Property

    ' Returns a collection of Tab objects that represent the tabs in the right/bottom side of a dual-display Lister. In a single-display Lister this will return an empty collection.
    Property Get tabsright ' Return Type collection:Tab
    End Property

    ' Returns the current title of the Lister window.
    Property Get title ' Return Type string
    End Property

    ' Returns a collection of Toolbar objects representing all currently open toolbars in this Lister.
    Property Get toolbars ' Return Type collection:Toolbar 
    End Property

    ' Lister window top-edge coordinate;
    Property Get top ' Return Type int
    End Property

    ' Indicates whether or not the folder tree is currently open. Possible values are:
    Property Get tree ' Return Type int
    End Property

    ' If the utility panel is currently open, returns a string indicating the currently selected utility page. Possible values are find, sync, dupe, undo, filelog, ftplog, otherlog, email.
    Property Get utilpage ' Return Type string
    End Property

    ' Indicates whether or not the utility panel is currently open. Possible values are:
    Property Get utilpane ' Return Type int
    End Property

    ' This Vars object represents all defined variables with Lister scope (that are scoped to this Lister).
    Property Get vars ' Return Type object:Vars 
    End Property

    ' Indicates whether or not the viewer pane is currently open. Possible values are:
    Property Get viewpane ' Return Type int
    End Property

    ' Creates a new Dialog object, that lets you display dialogs and popup menus. The dialog's window property will be automatically assigned to this Lister.
    Function Dlg ' Return Type object:Dialog 
    End Function

    ' Used to change how the lister window is grouped with other Opus windows on the taskbar. Specify a group name to move the window into an alternative group, or omit the group argument to reset back to the default group. If one or more windows are moved into the same group, they will be grouped together, separate from other the default group.
    Function SetTaskbarGroup(<string:group>) ' Return Type bool
    End Function

    ' The first time a script accesses a particular Lister object, a snapshot is taken of the Lister state. If the script then makes changes to that Lister (e.g. it opens a new tab, or moves the window), these changes will not be reflected by the object. To re-synchronize the object with the Lister, call the Lister.Update method.
    Sub Update
    End Sub
End Class

'The Listers object is a collection of all currently open Lister windows (each one represented by a Listerobject). It can be obtained from the DOpus.listers property.
Class Listers 'Default Return collection:Lister
    ' Returns a Lister object representing the most recently active Lister window.
    Property Get lastactive ' Return Type object:Lister
    End Property

    'The first time a script accesses the DOpus.listers property, a snapshot is taken of all currently open Listers. If the script then opens or closes Listers itself, these changes will not be reflected by this collection. To re-synchronize the collection, call the Update method.
    Sub Update
    End Sub
End Class

'This object is similar to an array or vector (e.g. Vector) in that it can store one or more objects, but has the advantage of using a dictionary system to locate objects rather than numeric indexes. It is obtained from the DOpusFactory.Map method.
Class Map
    ' Returns the number of elements the Map currently holds.
    Property Get count ' Return Type int
    End Property

    ' Returns True if the Map is empty, False if not.
    Property Get empty ' Return Type bool
    End Property

    ' A synonym for count.
    Property Get length ' Return Type int
    End Property

    ' A synonym for count.
    Property Get size ' Return Type int
    End Property

    ' Copies the contents of another Map to this one.
    Sub assign(<Map:from>)
    End Sub

    ' Clears the contents of the Map.
    Sub clear
    End Sub

    ' Erases the element matching the specified key, if it exists in the map.
    Sub erase(<variant:key>)
    End Sub

    ' Returns True if the specified key exists in the map.
    Function exists(<variant:key>) ' Return Type bool
    End Function

    ' Merges the contents of another Map with this one.
    Sub merge(<Map:from>)
    End Sub
End Class

'This object represents a file or folder's metadata. It can be obtained from the Item.metadata property, as well as the FSUtil.GetMetadata method.
Class Metadata ' Default Returns a string indicating the primary type of metadata available in this object. The string will be one of the following: none, video, audio, image, font, exe, doc, other.
    ' Returns an AudioMeta object providing access to audio metadata. The properties of this object are generally returned as their appropriate underlying type (e.g. a numeric field like ""track number"" will be returned as an int).
    Property Get audio ' Return Type object:AudioMeta 
    End Property

    ' Returns an AudioMeta object that provides access to the unmodified text form of the audio metadata. This provides access to the same text as displayed in a Lister. For example, a numeric field like ""track number"" would be returned as a string rather than an int.
    Property Get audio_text ' Return Type object:AudioMeta 
    End Property

    ' Returns a DocMeta object providing access to document metadata.
    Property Get doc ' Return Type object:DocMeta 
    End Property

    ' Returns a DocMeta object that provides access to the unmodified text form of the document metadata.
    Property Get doc_text ' Return Type object:DocMeta 
    End Property

    ' Returns an ExeMeta object providing access to executable (program) metadata.
    Property Get exe ' Return Type object:ExeMeta 
    End Property

    ' Returns an ExeMeta object that provides access to the unmodified text form of the program metadata.
    Property Get exe_text ' Return Type object:ExeMeta 
    End Property

    ' Returns a FontMeta object providing access to font file metadata.
    Property Get font ' Return Type object:FontMeta 
    End Property

    ' Returns an ImageMeta object providing access to picture metadata.
    Property Get image ' Return Type object:ImageMeta 
    End Property

    ' Returns an ImageMeta object that provides access to the unmodified text form of the picture metadata.
    Property Get image_text ' Return Type object:ImageMeta
    End Property

    ' Returns an OtherMeta object that provides access to miscellaneous metadata.
    Property Get other ' Return Type object:OtherMeta 
    End Property

    ' Returns a collection of strings corresponding to the tags that are assigned to this item.
    Property Get tags ' Return Type collection:string
    End Property

    ' Returns a VideoMeta object providing access to video metadata.
    Property Get video ' Return Type object:VideoMeta 
    End Property

    ' Returns a VideoMeta object that provides access to the unmodified text form of the video metadata.
    Property Get video_text ' Return Type object:VideoMeta
    End Property
End Class

'The Msg object represents a script dialog input event message. It’s returned by the Dialog.GetMsg method which you call when running the message loop for a detached dialog.
Class Msg ' Default Return Type bool. Returns True if the message is valid, or False if the dialog has been closed (which means you should exit your message loop).
    ' If the event type is checked, this indicates the check state of the item. If checkboxes are used in automatic mode, this will be the new check state of the item. In manual mode, this will indicate the existing state and it's up to you to change the state if desired.
    Property Get checked ' Return Type int
    End Property

    ' Returns the name of the control involved in the event. You can get a Control object representing the control by passing this string to the Dialog.Control method.
    Property Get control ' Return Type string
    End Property

    ' For resize events, this property returns the new width of the dialog.
    Property Get cx ' Return Type int
    End Property

    ' For resize events, this property returns the new height of the dialog.
    Property Get cy ' Return Type int
    End Property

    ' If the event type is focus, indicates the new focus state of the control - True if the control has gained the focus, or False if it's lost it.
    Property Get data ' Return Type int or bool
    End Property

    ' Returns the name of the parent dialog.
    Property Get dialog ' Return Type string
    End Property

    ' Returns a string indicating the event that occurred. Currently defined events are:
    Property Get event ' Return Type string
    End Property

    ' Returns True if the control had focus when the message was generated.
    Property Get focus ' Return Type bool
    End Property

    ' Returns the current selection index for a combo box, list box or tab control.
    Property Get index ' Return Type int
    End Property

    ' Returns the horizontal position of the mouse cursor when the message was generated.
    Property Get mousex ' Return Type int
    End Property

    ' Returns the vertical position of the mouse cursor when the message was generated.
    Property Get mousey ' Return Type int
    End Property

    ' For a drop event, this property returns a Vector of Item objects, representing the files that were dropped onto your dialog.
    Property Get object ' Return Type variant
    End Property

    ' Returns a string indicating the qualifier keys (if any) that were held down when the message was generated.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns True if the message is valid, or False if the dialog has been closed.
    Property Get result ' Return Type bool
    End Property

    ' For a dialog tab control, returns the name of the parent tab (if the control is on a dialog that's inside a tab control).
    Property Get tab ' Return Type string
    End Property

    ' For the dblclk, editchange and selchange events, returns the current contents of the edit field (or selected item label).
    Property Get value ' Return Type string
    End Property
End Class

'This object provides general metadata properties relating to files and folders. It is obtained from the Metadata object.
Class OtherMeta
    ' An automatically generated description string for the item. This is the same string that is shown in the Description column in a Lister. Opus automatically generates the description for various types of files using the other metadata in ways that make the most sense.
    Property Get autodesc ' Return Type string
    End Property

    ' For a folder, the size of which has been calculated via GetSizes or similar, this provides the number of sub-folders directly underneath the folder.
    Property Get dircount ' Return Type int
    End Property

    ' Similar to dircount, this provides the total number of sub-folders underneath the folder (this is a recursive count - it includes sub-sub-folders, sub-sub-sub-folders, etc.)
    Property Get dircounttotal ' Return Type int
    End Property

    ' For a folder, the size of which has been calculated via GetSizes or similar, this provides the number of files directly located in that folder.
    Property Get filecount ' Return Type int
    End Property

    ' Similar to filecount, this provides the total number of files in the folder and all its sub-folders, sub-sub-folders, etc.
    Property Get filecounttotal ' Return Type int
    End Property

    ' For a folder, the size of which has been calculated via GetSizes or similar, this returns a string giving a summary of the contents of the folder.
    Property Get foldercontents ' Return Type string
    End Property

    ' A description automatically generated for the item by its parent virtual file system.
    Property Get nsdesc ' Return Type string
    End Property

    ' Returns the user-assigned rating for this file or folder.
    Property Get rating ' Return Type int
    End Property

    ' Returns a Path object representing the target path of shortcuts and links.
    Property Get target ' Return Type object:Path 
    End Property

    ' Returns a string indicating the type of the link (unknown, linkfile, dosfile, url, junction, softlink).
    Property Get target_type ' Return Type string
    End Property

    ' Returns the user-assigned description for the file or folder.
    Property Get usercomment ' Return Type string
    End Property
End Class

'This object represents a file system path. It contains several methods to manipulate the path. Path objects are returned by several properties and can be created by the FSUtil.NewPath method.
Class Path ' Default Returns the full path as a string.
    ' Returns the filename part of the path (the last component).
    Property Get filepart ' Return Type string
    End Property

    ' If this object represents a short pathname, this property returns the ""long"" equivalent.
    Property Get longpath ' Return Type object:Path
    End Property

    ' Returns the path minus the last component.
    Property Get pathpart ' Return Type string
    End Property

    ' If this object represents a long pathname, this property returns the ""short"" equivalent, if it has one. Note that short paths are disabled by default in Windows 10.
    Property Get shortpath ' Return Type object:Path
    End Property

    ' Returns the filename stem of the path (i.e. filepart minus ext).
    Property Get stem ' Return Type string
    End Property

    ' Returns the filename stem taking multi-part extensions into account. For example, stem might return ""pictures.part1"" whereas stem_m would return ""pictures"".
    Property Get stem_m ' Return Type string
    End Property

    ' Returns True if a call to the Parent method would succeed.
    Property Get test_parent ' Return Type bool
    End Property

    ' Returns True if a call to the Root method would succeed.
    Property Get test_root ' Return Type bool
    End Property

    ' Adds the specified name to the path (it will become the last component). As well as a string, you can pass a Vector of strings and all items in the vector will be added to the path.
    Sub Add(<string:name> Or <Vector:string>)
    End Sub

    ' Removes the last component of the path. Returns False if the path does not have a valid parent.
    Function Parent ' Return Type bool
    End Function

    ' Compares the beginning of the path with the ""old"" string, and if it matches replaces it with the ""new"" string. The match is performed at the path component level - for example, an ""old"" string of ""C:\Foo"" would match the path ""C:\Foo\Bar"" but not ""C:\FooBar"". If the optional ""wholepath"" argument is set to True then the whole path must match rather than just its beginning. Returns True if the string matched the path or False otherwise.
    Function ReplaceStart(<string:old>,<string:New>,<bool:wholepath>) ' Return Type bool
    End Function

    ' Strips off all but the first component of the path. Returns False if the path is already at the root.
    Function Root ' Return Type bool
    End Function

    ' Sets the path represented by the Path object to the specified string. You can also set one Path object to the value of another. If you pass a Vector of strings the path will be built from the items in the vector.
    Sub Set(<string:path> Or <Path:path> Or <Vector:string>)
    End Sub

    ' Returns a Vector of strings representing the components of the path. For example, if the path is C:\Foo\Bar, the vector will contain three items - ""C:\"", ""Foo"" and ""Bar"". By default all components of the path are returned, but you can optionally provide the index of the first component and also the number of components to return.
    Function Split(<int:first>,<int:count>) ' Return Type Vector:string
    End Function




    ' Returns the number of components in the path.
    Property Get components ' Return Type int
    End Property

    ' Returns a Vector of ints representing the physical disk drive or drives that this path resides on.
    Property Get disks ' Return Type Vector:int
    End Property

    ' Returns the drive number the path refers to (1=A, 2=B, etc.) or 0 if the path does not specify a drive. You can also change the drive letter of the path (while leaving the following path components alone) by modifying this value.
    Property Get drive ' Return Type int
    End Property

    ' Returns the filename extension of the path (the sub-string extending from the last . in the final component to the end of the string). This method does not check if the path actually refers to a file.
    Property Get ext ' Return Type string
    End Property

    ' Returns the filename extension of the path, taking multi-part extensions into account. For example, ext might return "".rar"" whereas ext_m would return "".part1.rar"".
    Property Get  ext_m ' Return Type string
    End Property

End Class

'This object represents a progress dialog, that lets you visually indicate to the user the progress of your script function. It is obtained from the Command.progress property.
Class Progress
    ' Before calling Init, set to True if the Abort button should be available, or False to disable it.
    Property Get abort ' Return Type bool
    End Property

    ' Before calling Init, set to True if the dialog should show progress in bytes rather than whole files.
    Property Get bytes ' Return Type bool
    End Property

    ' Before calling Init, set to True if the dialog should delay before appearing after the Show method is called. The delay is configured by the user in Preferences.
    Property Get delay ' Return Type bool 
    End Property

    ' Before calling Init, set to True to enable a ""full size"" progress indicator with two separate progress bars (one for files and one for bytes).
    Property Get full ' Return Type bool
    End Property

    ' Before calling Init, set to True if the dialog should be owned by its parent window (the parent is given later, when the dialog is created via the Init method).
    Property Get owned ' Return Type bool
    End Property

    ' Before calling Init, set to True if the Pause button should be available.
    Property Get pause ' Return Type bool
    End Property

    ' Before calling Init, set to True if the Skip button should be available. (This just makes it so the Skip button can be enabled. You must still call EnableSkip later to actually enable it; usually once per file.)
    Property Get skip ' Return Type bool
    End Property

    ' Adds the specified number of files to the operation total. The bytes argument is optional - in a ""full size"" progress indicator this lets you add to the total byte size of the operation.
    Sub AddFiles(<int:count>,<FileSize:bytes>)
    End Sub

    ' Clears the state of the three ""control"" buttons (Abort / Pause / Skip) so they no longer register as being clicked when GetAbortState is called.
    Sub ClearAbortState
    End Sub

    ' Enables the progress dialog's Skip button. For EnableSkip to work, you must have set the skip property to True before the progress dialog was created by the Init method.
    Sub EnableSkip(<bool:enable>,<bool:delay>,<bool:clear>)
    End Sub

    ' Finish the current file. If the byte size of the current file has been set the total progress will be advanced by any remaining bytes.
    Sub FinishFile
    End Sub

    ' Polls the state of the three ""control"" buttons. This returns a string that indicates which, if any, of the three buttons have been clicked by the user. The button states are represented by the following letters in the returned string:
    Function GetAbortState(<bool:autoPause>,<string:wanted>,<bool:simple>) ' Return Type string
    End Function

    ' Hides the progress indicator dialog. The dialog object itself remains valid, and can be redisplayed with the Show method if desired.
    Sub Hide
    End Sub

    ' Hides or shows the ""XX bytes / YY bytes"" string in the progress dialog. You can use this to hide the string if the progress does not indicate a number of bytes (e.g. when it indicates a percentage). Pass True for the show argument to show the string and False to hide it.
    Sub HideFileByteCounts(<bool:show>)
    End Sub

    ' Initializes the dialog. This method causes the actual dialog to be created, although it will not be displayed until the Show method is called. The fundamental properties shown above must be set before this method is called - once the dialog has been created they can not be altered.
    Sub Init(<Tab:parent>,or <Lister:parent>,<string:title>)
    End Sub

    ' Resets the byte count for the current file to zero.
    Sub InitFileSize
    End Sub

    ' Resets the total completed file and byte counts to zero.
    Sub Restart
    End Sub

    ' Sets the total completed byte count.
    Sub SetBytesProgress(<FileSize:bytes>)
    End Sub

    ' Sets the size of the current file.
    Sub SetFileSize(<FileSize:bytes>)
    End Sub

    ' Sets the total number of files.
    Sub SetFiles(<int:count>)
    End Sub

    ' Sets the total completed file count.
    Sub SetFilesProgress(<int:count>)
    End Sub

    ' Sets the text at the top of the dialog that indicates the source and destination of an operation. The header argument refers to the string that normally says From: - this allows you to change it in case that term is not applicable to your action. The from argument is the source path, and the to argument (if there is one) is the destination path. Note that if you specify a destination path this always has a To: header appended to it.
    Sub SetFromTo(<string:header>,<string:from>,<string:to>)
    End Sub

    ' Sets the name of the current file.
    Sub SetName(<string:name>)
    End Sub

    ' Sets the current progress as a percentage (from 0 to 100).
    Sub SetPercentProgress(<int:percent>)
    End Sub

    ' Sets the text displayed in the status line at the top of the dialog.
    Sub SetStatus(<string:status>)
    End Sub

    ' Sets the title of the dialog.
    Sub SetTitle(<string:title>)
    End Sub

    ' Sets the type of the current item - either file or dir.
    Sub SetType(<string:type>)
    End Sub

    ' Displays the progress indicator dialog. Call this once you have created the dialog using the Init method.
    Sub Show
    End Sub

    ' Skips over the current file. Set the complete argument to True to have the file counted as ""complete"", or False to count it as ""skipped"".
    Sub SkipFile(<bool:complete>)
    End Sub

    ' Step the byte progress indicator the specified number of bytes.
    Sub StepBytes(<FileSize:bytes>)
    End Sub

    ' Step the file progress indicator the specified number of files.
    Sub StepFiles(<int:count>)
    End Sub
End Class

'This object provides information about the state of the quick filter in a tab. It's obtained from the Tab.quickfilter property.
Class QuickFilter ' Default Returns the current filter string, if any.
    ' Returns True if the auto-clear mode is set in Preferences.
    Property Get autoclear ' Return Type bool
    End Property

    ' Returns True if the auto-star mode is set in Preferences.
    Property Get autostar ' Return Type bool 
    End Property

    ' Returns True if the filter is disabled.
    Property Get disable ' Return Type bool
    End Property

    ' Returns True if easy mode is selected.
    Property Get easymode ' Return Type bool
    End Property

    ' Returns the current filter string.
    Property Get filter ' Return Type string
    End Property

    ' Returns True if all folders are being hidden.
    Property Get hidealldirs ' Return Type bool
    End Property

    ' Returns True if all files are being hidden.
    Property Get hideallfiles ' Return Type bool
    End Property

    ' Returns True if filtering in flatview is enabled.
    Property Get overrideflatview  ' Return Type bool
    End Property

    ' Returns True if partial matching is enabled.
    Property Get partial ' Return Type bool
    End Property

    ' Returns True if realtime filtering is enabled.
    Property Get realtime ' Return Type bool
    End Property

    ' Returns True if regular expression mode is enabled.
    Property Get regex ' Return Type bool
    End Property

    ' Returns True if all folders are being shown.
    Property Get showalldirs ' Return Type bool
    End Property

    ' Returns True if all files are being shown.
    Property Get showallfiles ' Return Type bool
    End Property

    ' Returns True if Show Everything mode is on, which overrides (almost) all filtering.
    Property Get showeverything ' Return Type bool
    End Property
End Class

'The Rect object represents a rectangle.
Class Rect
    ' Returns the left edge of the rectangle.
    Property Get left ' Return Type int
    End Property

    ' Returns the top edge of the rectangle.
    Property Get top ' Return Type int
    End Property

    ' Returns the right edge of the rectangle.
    Property Get right ' Return Type int
    End Property

    ' Returns the bottom edge of the rectangle.
    Property Get bottom  ' Return Type int
    End Property

    ' Returns the width of the rectangle. Equal to right-left.
    Property Get width ' Return Type int
    End Property

    ' Returns the height of the rectangle. Equal to bottom-top.
    Property Get height ' Return Type int
    End Property

    ' Returns a string describing the rectangle's position and size, as a convenience when debugging scripts. The format is ""(L,T - R,B; WxH)"" i.e. Left, Top, Right, Bottom, Width, and Height.
    Function ToString ' Return Type <string>
    End Function
End Class

'This object represents the results of a command (the error code in the case of failure, plus any new tabs or Listers created by the command). It is obtained from the Command.results property.
Class Results
    ' Indicates whether or not the command ran successfully. Zero indicates the command could not be run or was aborted; any other number indicates the command was run for at least some files. (Note that this is not the ""exit code"" for external commands. For external commands it only indicates whether or not Opus launched the command. If you need the exit code of an external command, use the WScript.Shell Run or Exec methods to run the command.)
    Property Get result ' Return Type int
    End Property

    ' This property returns a collection of Tab objects representing any new tabs created by the command.
    Property Get newtabs ' Return Type collection:Tab 
    End Property

    ' This property returns a collection of Lister objects representing any new Listers created by the command.
    Property Get newlisters  ' Return Type collection:Lister 
    End Property

    ' This property returns a collection of Viewer objects representing any new image viewers created by the command. (This is only for standalone viewers, not the viewer pane.)
    Property Get newviewers  ' Return Type collection:Viewer 
    End Property
End Class

'This object represents a script-defined column. It is obtained from the ScriptInitData.AddColumn method, while processing the OnInit event.
Class ScriptColumn
    ' If this is set to True (which is the default), and the file display is grouped by this column, Opus will generate the groups automatically based on the column value. If you set this to False, Opus will expect you to provide grouping information in your OnScriptColumn function.
    Property Get autogroup ' Return Type bool
    End Property

    ' Set to True (or 1) to force Opus to update the value for this column when a file changes. You can also set this value to 2 to force Opus to update the value when the file's attributes change (normally it would only update if the file modification time or size changed).
    Property Get autorefresh ' Return Type bool or int
    End Property

    ' This property lets you control the default sort behavior for your column. Normally when the user clicks the column header to sort by a column the column is initially sorted in ascending order, and then clicking again reverses the sort order. If you set defsort to -1, the first click on the column header will sort in descending order. Date and size fields have this behavior set by default.
    Property Get defsort ' Return Type int
    End Property

    ' Specifies a default width for your column, which will be used unless the file display has auto-sizing enabled. If you specify a simple integer value this represents a width measured in average characters (e.g. 12 specifies 12 average characters wide). You can also specify an absolute number of pixels by adding the px suffix (e.g. ""150px"" specifies 150 pixels).
    Property Get defwidth ' Return Type int or string
    End Property

    ' For graph columns, specifies the first graph color set. The graph will be displayed in these colors as long as its percentage is below the threshold.
    Property Get graph_colors ' Return Type object:Vector 
    End Property

    ' Similar to graph_colors, this property lets you configure a second set of colors for a graph column that will be used when the graph value exceeds the threshold.
    Property Get graph_colors2 ' Return Type object:Vector 
    End Property

    ' For graph columns, specifies the percentage threshold at which the graph will switch from the first color set to the second (e.g. a blue graph goes red to indicate a drive is nearly full). Set the threshold to -1 to disable the second color set altogether.
    Property Get graph_threshold ' Return Type
    End Property

    ' If the autogroup property is set to False, the grouporder property lets you control the order your column's groups appear in. Each group should be listed in the string in the desired order, separated by a semi-colon (e.g. ""Never Modified;Modified""). If not provided, groups will default to sorting alphabetically.
    Property Get grouporder ' Return Type string
    End Property

    ' If this property is set, this defines the string that will be displayed in the column header when this column is added to a Lister. If not set, the label value will be used.
    Property Get header ' Return Type string
    End Property

    ' Set this to True if you  want your column to be only available for use in Info Tips. You might want this if your column takes a significant amount of time to return a value, in which case the user would probably only want to use it in an Info Tip so they can see the value on demand. If set to False (the default) the column will be available everywhere.
    Property Get infotiponly ' Return Type bool
    End Property

    ' This field lets you control the justification of your column. If not specified, columns default to left justify. Acceptable values are center, left, right and path.
    Property Get justify ' Return Type string
    End Property

    ' If this is set to True, and the user has the Sort-field specific key scrolling Preferences option enabled, then your column will participate in this special mode.
    Property Get keyscroll ' Return Type bool
    End Property

    ' Use this to set a label for the column. This is displayed in the column header when the column is added to a Details/Power mode file display (unless overridden by the header property), and in various column lists such as in the Folder Options dialog.
    Property Get label ' Return Type string
    End Property

    ' If you add strings to this Vector (e.g. via the push_back method) it will be used to provide a drop-down list of possible values when searching on this column using the Advanced Find function.
    Property Get match ' Return Type Vector:string
    End Property

    ' If the column type is set to stars this property lets you specify the maximum number of stars that will be used. This is used to ensure the column is sized correctly.
    Property Get maxstars ' Return Type int
    End Property

    ' This is the name of the method in your script that provides the actual values for your new column. This would typically be set to OnXXXXX where XXXXX is the name of the command, however any method name can be used.
    Property Get method ' Return Type string
    End Property

    ' If your script implements multiple columns that require common calculations to perform, you may wish to set the multicol property. If this is set to True then your column handler function has the option of returning data for multiple columns simultaneously, rather than just the specific column it is being invoked for.
    Property Get multicol ' Return Type bool
    End Property

    ' This is the raw name of the column. This determines the name that can be used to control the column programmatically (for example, the Set COLUMNSTOGGLE command can be used to toggle a column on or off by name).
    Property Get name ' Return Type string
    End Property

    ' Set to True to force Opus to update the value for this column when a file's name changes.
    Property Get namerefresh ' Return Type bool
    End Property

    ' Set to True to prevent the file display being grouped by this column.
    Property Get nogroup ' Return Type bool
    End Property

    ' Set to True to prevent the file display being sorted by this column.
    Property Get nosort ' Return Type bool
    End Property

    ' Time, in milliseconds, before Opus may give up waiting for calculation of a column value.
    Property Get timeout ' Return Type int
    End Property

    ' This field lets you set the default type of the column.
    Property Get type ' Return Type string
    End Property

    ' Allows you to associate a data value with a column. The value will be passed to your column handler in the ScriptColumnData.userdata property
    Property Get userdata ' Return Type variant
    End Property
End Class

'This object represents a script-defined internal command. It is obtained from the ScriptInitData.AddCommand method, while processing the OnInit event.
Class ScriptCommand
    ' Use this to set a description for the command, that is displayed in the Customize dialog when the user selects the command from the Commands tab.
    Property Get desc ' Return Type string
    End Property

    ' Set to True to hide this command from the drop-down command list shown in the command editor. This lets you add commands that can still be used in buttons and hotkeys but won't clutter up the command list.
    Property Get hide ' Return Type bool
    End Property

    ' Use this property to assign a default icon to this command. You can specify the name of an internal icon (if you want to specify an icon from a particular set, use setname:iconname - use this if you have bundled your script in a script package with its own icon set) or the path of an external icon or image file.
    Property Get icon ' Return Type string
    End Property

    ' Use this to set a label for the command. This is displayed in the Commands tab of the Customize dialog (under the Script Commands category), and will form the default label of the button created if the user drags that command out to a toolbar.
    Property Get label ' Return Type string
    End Property

    ' This is the name of the method that Opus will call in your script when the command is invoked. This would typically be set to OnXXXXX where XXXXX is the name of the command, however any method name can be used.
    Property Get method ' Return Type string
    End Property

    ' This is the name of the command. This determines the name that will invoke the command when it is used in buttons and hotkeys.
    Property Get name ' Return Type string
    End Property

    ' This lets you specify an optional command line template for the command.
    Property Get template ' Return Type string
    End Property
End Class

'This object represents script-defined configuration data that Opus stores for each script. The configuration items are initialised via the ScriptInitData.config property, and are then available to the script via the Script.config property.
Class ScriptConfig ' The properties of the ScriptConfig object are entirely determined by the script itself.
End Class

'The ScriptStrings object is returned by the DOpus.strings property. It lets you access any strings defined via string resources.
Class ScriptStrings
    ' Returns a Vector of strings representing the languages that strings have been defined for.
    Property Get langs ' Return Type object:Vector 
    End Property

    ' Returns the text of a string specified by name. The name must match the name used in the string resources.
    Function Get(<string:name>,<string:language>) ' Return Type string
    End Function

    ' Returns True if strings in the specified language are defined in the resources.
    Function HasLanguage(<string:language>) ' Return Type bool
    End Function
End Class

'The ShellProperty object represents a shell property - an item of metadata for a file or folder that comes from Windows or third-party extensions. The FSUtil.GetShellPropertyList method lets you retrieve a list of available shell properties.
Class ShellProperty
    ' The default width in pixels a column displaying this property should use.
    Property Get defwidth ' Return Type int
    End Property

    ' The display name of this property (the name that should be shown to users).
    Property Get display_name ' Return Type string
    End Property

    ' The default column justification for this property (left, right, center).
    Property Get justify ' Return Type string
    End Property

    ' The PKEY (property key) for this property. This is a property's unique ID and the canonical way to refer to a property. You can use the raw_name and display_name values to access properties as well, but they are potentially inaccurate (since it's possible to have two properties with the same name) and also slower as the property has to be looked up by name each time.
    Property Get pkey ' Return Type string
    End Property

    ' An internal name used by the property provider.
    Property Get raw_name ' Return Type string
    End Property

    ' The type of data this property returns; string, number, datetime are the only supported types currently.
    Property Get type ' Return Type string
    End Property
End Class

'A SmartFavorite object represents an entry for a folder in the SmartFavorites table. It is retrieved by enumerating or indexing the SmartFavorites object.
Class SmartFavorite
    ' Returns the path this entry represents, as a Path object.
    Property Get path ' Return Type object:Path 
    End Property

    ' Returns the number of points this entry has as a source folder. The point score is used by Opus to determine which folders to display.
    Property Get points ' Return Type int
    End Property

    ' Returns the number of points this entry has as a destination folder.
    Property Get destpoints ' Return Type int
    End Property
End Class

'The SmartFavorites object lets you query the contents of the SmartFavorites table. It is retrieved from the DOpus.smartfavorites property.
Class SmartFavorites ' Default Return Type collection:SmartFavorite
    ' Returns the number of points an entry must have before it would be displayed in the SmartFavorites list.
    Property Get threshhold ' Return Type int
    End Property

    ' Returns the maximum number of entries that would be displayed in the SmartFavorites list.
    Property Get max ' Return Type int
    End Property
End Class

'The SortOrder object is returned by the Format.manual_sort_order property if manual sort mode is active. It lets you query and modify the sort order.
Class SortOrder
    ' Returns a Vector of strings representing the current sort order of files in the folder. If multiple manual sort orders have been defined, you can provide the name of a specific sort order as an argument to this method. If called with no arguments it returns the current sort order by default.
    Function GetOrder(<string:name>) ' Return Type object:Vector 
    End Function

    ' You can pass this method a Vector of strings to change the sort order of the current folder. You can optionally provide the name of a sort order as the second parameter if you’ve got more than one sort order defined.
    Sub SetOrder(<Vector:order>,<string:name>)
    End Sub

    ' Resets the manual sort order to the currently selected sort order (e.g. if the file display header indicates that it is sorted by name, ResetOrder would reset to filename order). You can optionally provide the name of a sort order as the second parameter if you’ve got more than one sort order defined.
    Sub ResetOrder(<string:name>)
    End Sub
End Class

'This object is similar to an array or vector (e.g. Vector) of strings, but has the advantage of using a dictionary system to locate strings rather than numeric indexes. It is obtained from the DOpusFactory.StringSet and StringSetI methods.
Class StringSet
    ' Returns the number of elements the StringSet currently holds.
    Property Get count ' Return Type int
    End Property

    ' Returns True if the StringSet is empty, False if not.
    Property Get empty ' Return Type bool
    End Property

    ' A synonym for count.
    Property Get length ' Return Type int
    End Property

    ' A synonym for count.
    Property Get size ' Return Type int
    End Property

    ' Copies the contents of another StringSet to this one. You can also pass an array of strings or Vector object.
    Sub assign(<StringSet:from>)
    End Sub

    ' Clears the contents of the StringSet.
    Sub clear
    End Sub

    ' Erases the string if it exists in the set.
    Sub erase(<string>)
    End Sub

    ' Returns True if the specified string exists in the set.
    Function exists(<string>) ' Return Type bool
    End Function

    ' Inserts the string into the set if it doesn't already exist. Returns True if successful.
    Function insert(<string>) ' Return Type bool
    End Function

    ' Merges the contents of another StringSet with this one.
    Sub merge(<StringSet:from>)
    End Sub
End Class

'This object provides utility functions for string encoding and decoding. It is obtained from the DOpusFactory.StringTools method.
Class StringTools
    ' Decodes an encoded string or data.
    Function Decode(<Blob:source> or <string:source>,<string:format>) ' Return Type string or Blob 
    End Function

    ' Encodes a string or data.
    Function Encode(<Blob:source> or <string:source>,<string:format>) ' Return Type string or Blob 
    End Function

    ' Tests the input string to see if it only contains characters that can be represented in ASCII.
    Function IsASCII(<string:input>) ' Return Type bool
    End Function
End Class

'The SysInfo object is created by the DOpusFactory.SysInfo method. It lets scripts access miscellaneous system information that may not be otherwise easy to obtain from a script.
Class SysInfo
    ' Allows you to test if a named process is currently running, and returns the process's ID if so. If the process isn't running 0 is returned. You can use wildcards or (by prefixing the pattern with regex:) regular expressions.
    Function FindProcess(string) ' Return Type int
    End Function

    ' If called with no arguments, returns a Vector of Rect objects which provide information about the positions and sizes of the display monitors in the system.
    Function Monitors(none or int:index) ' Return Type Vector:Rect 
    End Function

    ' Returns the index of the monitor the mouse pointer is currently positioned on.
    Function MouseMonitor ' Return Type int
    End Function

    ' Returns the current x-coordinate of the mouse pointer.
    Function MousePosX ' Return Type int
    End Function

    ' Returns the current y-coordinate of the mouse pointer.
    Function MousePosY ' Return Type int
    End Function

    ' Returns a Rect giving the size of the invisible border around windows.
    Function ShadowBorder ' Return Type Rect
    End Function

    ' Similar to the Monitors method, documented above, except it returns the work area of each monitor rather than the full monitor area.
    Function WorkAreas(none or int:index) ' Return Type Vector:Rect or Rect 
    End Function
End Class

'This object represents a folder tab in a Lister. A Lister's tabs are available via various Lister object properties (e.g. Lister.activetab) and also used to specify the source/destination of a command (e.g. Command.sourcetab).
Class Tab
    ' Returns a collection of Item objects that represents all the files and folders currently displayed in this tab.
    Property Get all ' Return Type collection:Item 
    End Property

    ' Returns a collection of Path objects that represents the paths in the ""backward"" history list for this tab (i.e. the folders you would get to by clicking the Back button).
    Property Get backlist ' Return Type collection:Path 
    End Property

    ' Returns the tab's assigned color (if one has been assigned via, for example, the Go TABCOLOR command). The color is returned as a string in R,G,B format.
    Property Get color ' Return Type string
    End Property

    ' Returns the current path from the tab's breadcrumb control (if it has one), including any ghost path.
    Property Get crumbpath ' Return Type object:Path 
    End Property

    ' Returns a collection of Item objects that represents all the folders currently displayed in this tab.
    Property Get dirs ' Return Type collection:Item 
    End Property

    ' Returns True if the tab is marked as dirty, indicating its list of contents may be out of date. This can happen if the tab is in the background and the user has turned off the Preferences / Folder Tabs / Options / Process file changes in background tabs option.
    Property Get dirty ' Return Type bool
    End Property

    ' Returns the currently displayed label of this tab.
    Property Get displayed_label ' Return Type string
    End Property

    ' Returns a collection of FileGroup objects that represents all the file groups in the tab (when the tab is grouped). You can use the format.group_by property to test if the tab is grouped or not.
    Property Get filegroups ' Return Type collection:FileGroup 
    End Property

    ' Returns a collection of Item objects that represents all the files currently displayed in this tab.
    Property Get files ' Return Type collection:Item
    End Property

    ' Returns a Format object representing the current folder format in this tab.
    Property Get format ' Return Type object:Format 
    End Property

    ' Returns a collection of Path objects that represents the paths in the ""forward"" history list for this tab (i.e. the folders you would get to by clicking the Forward button).
    Property Get forwardlist ' Return Type collection:Path
    End Property

    ' Returns a collection of Item objects that represents all the files and folders currently hidden from this tab.
    Property Get hidden ' Return Type collection:Item
    End Property

    ' Returns a collection of Item objects that represents all the folders currently hidden from this tab.
    Property Get hidden_dirs ' Return Type collection:Item
    End Property

    ' Returns a collection of Item objects that represents all the files currently hidden from this tab
    Property Get hidden_files ' Return Type collection:Item
    End Property

    ' Returns the current assigned tab label. Note that this may be an empty string if no custom label has been assigned. The displayed_label property returns the currently displayed label in all cases.
    Property Get label ' Return Type string
    End Property

    ' If this tab is linked to another tab, returns a Tab object representing the linked tab. If this tab is not linked this property returns 0.
    Property Get linktab ' Return Type object:Tab 
    End Property

    ' Returns a Lister object representing the parent Lister that owns this tab.
    Property Get lister ' Return Type object:Lister 
    End Property

    ' Returns the current lock state of the tab; one of ""off"", ""on"", ""changes"", ""reuse"".
    Property Get lock ' Return Type string
    End Property

    ' Returns True if this tab is linked in Navigation Lock mode. This property does not exist if the tab is not linked, so make sure you check the value of linktab first.
    Property Get navlock ' Return Type bool
    End Property

    ' Returns the current path shown in this tab.
    Property Get path ' Return Type object:Path 
    End Property

    ' Returns a QuickFilter object providing information about the state of the quick filter in this tab.
    Property Get quickfilter ' Return Type object:QuickFilter 
    End Property

    ' Returns True if this tab is currently on the right or bottom side of a dual-display Lister, and False otherwise.
    Property Get right ' Return Type bool
    End Property

    ' Returns a collection of Item objects that represents all the selected files and folders currently displayed in this tab. Note that if checkbox mode is turned on in the tab, this will be a collection of checked items rather than selected.
    Property Get selected ' Return Type collection:Item
    End Property

    ' Returns a collection of Item objects that represents all the selected folders currently displayed in this tab.
    Property Get selected_dirs ' Return Type collection:Item
    End Property

    ' Returns a collection of Item objects that represents all the selected files currently displayed in this tab
    Property Get selected_files ' Return Type collection:Item
    End Property

    ' Returns a TabStats object that provides various information about the tab, including the number of files, number of selected files, total size of selected files, etc. The ""selected"" counts provided by this object take checkbox mode into account (that is, if checkbox mode is currently turned on, the counts will be for checked files rather than for selected files).
    Property Get selstats ' Return Type object:TabStats 
    End Property

    ' Returns True if this tab is currently the source, and False otherwise.
    Property Get source ' Return Type bool
    End Property

    ' Returns a TabStats object that provides various information about the tab, including the number of files, number of selected files, total size of selected files, etc. Unlike selstats, this object does not take checkbox mode into account (so the ""selected"" counts will refer to selected rather than checked files).
    Property Get stats ' Return Type object:TabStats
    End Property

    ' This Vars object represents all defined variables with tab scope (that are scoped to this tab).
    Property Get vars ' Return Type object:Vars 
    End Property

    ' Returns True if this tab is currently visible (i.e. it is the active tab in either file display), and False otherwise.
    Property Get visible ' Return Type bool
    End Property

    ' Creates a new Dialog object, that lets you display dialogs and popup menus. The dialog's window property will be automatically assigned to this tab.
    Function Dlg ' Return Type object:Dialog
    End Function

    ' Returns an Item object representing the file or folder which has focus in the tab.
    Function GetFocusItem ' Return Type object:Item 
    End Function

    ' When a script accesses particular properties of a Tab object, a snapshot is taken of the tab's state.
    Sub Update
    End Sub
End Class

'This object represents a folder tab group. It's accessed by enumerating the TabGroups object.
Class TabGroup
    ' True if the Close existing folder tabs when opening this group option is turned on for this group. Only present when the folder property is False.
    Property Get closeexisting ' Return Type bool
    End Property

    ' The description of this tab group, if any. Only present when the folder property is False.
    Property Get desc ' Return Type string
    End Property

    ' True if the Define tabs on specific sides of a dual-display Lister option is turned on for this group. Only present when the folder property is False.
    Property Get dual ' Return Type bool
    End Property

    ' True if this object represents a folder within the tab group list, False if it's an actual tab group.
    Property Get folder ' Return Type bool
    End Property

    ' True if this tab group or folder should be hidden from menus which list tab groups. The group will still always be visible in Preferences.
    Property Get hidden ' Return Type bool
    End Property

    ' Returns a TabGroupTabList object representing the tabs in this group that open in the left/top side of a dual-display Lister. Only present when the folder property is False and the dual property is True.
    Property Get lefttabs ' Return Type object:TabGroupTabList 
    End Property

    ' The name of this group or folder.
    Property Get name ' Return Type string
    End Property

    ' Returns a TabGroupTabList object representing the tabs in this group that open in the right/bottom side of a dual-display Lister. Only present when the folder property is False and the dual property is True.
    Property Get righttabs ' Return Type object:TabGroupTabList
    End Property

    ' Returns a TabGroupTabList object representing the tabs in this group. Only present when both the folder and dual properties are False.
    Property Get tabs ' Return Type object:TabGroupTabList
    End Property

    ' Adds a new sub-folder to this tab group folder. Only available when the folder property is True. You can either provide a TabGroup object (which itself has the folder property set to True) or the name for the new folder. If the operation succeeds a TabGroup object is returned which represents the new folder. If the operation fails False is returned.
    Function AddChildFolder(<object:TabGroup> or <string:name>) ' Return Type object:TabGroup
    End Function

    ' Adds a new tab group to this tab group folder. Only available when the folder property is True. You can either provide a TabGroup object or the name for the new group. If the operation succeeds a TabGroup object is returned which represents the new tab group. If the operation fails False is returned.
    Function AddChildGroup (<object:TabGroup> or <string:name>) ' Return Type object:TabGroup
    End Function

    ' Deletes the child item (folder or tab group).
    Sub DeleteChild(<object:TabGroup>)
    End Sub

    ' Returns a duplicate of this tab group or folder. When it's returned the duplicate has not yet been added to a tab list.
    Function Duplicate ' Return Type object:TabGroup
    End Function

    ' In a tab group that has specific left and right tabs specified, this method links together a tab from the left side and a tab from the right side. Only available if the dual property is set to True. You can provide TabGroupTabEntry objects or the index numbers of the tabs you want to link.
    Sub Link(<object:TabGroupTabEntry>,<object:TabGroupTabEntry>,<string:type>)
    End Sub

    ' Unlinks the specified tab from its partner. Only available if the dual property is set to True.
    Sub Unlink(<object:TabGroupTabEntry> )
    End Sub
End Class

'This object provides access to and lets you modify the configured list of folder tab groups. It's obtained from the DOpus.tabgroups property.
Class TabGroups
    ' Adds a new folder to the list of tab groups. You can either provide a TabGroup object (which has the folder property set to True) or the name for the new folder. If the operation succeeds a TabGroup object is returned which represents the new folder. If the operation fails False is returned.
    Function AddChildFolder(<object:TabGroup> or <string:name>) ' Return Type object:TabGroup
    End Function

    ' Adds a new tab group to the list of tab groups. You can either provide a TabGroup object or the name for the new group. If the operation succeeds a TabGroup object is returned which represents the new tab group. If the operation fails False is returned.
    Function AddChildGroup(<object:TabGroup> or <string:name>) ' Return Type object:TabGroup
    End Function

    ' Deletes the child item (folder or tab group).
    Sub DeleteChild(<object:TabGroup>)
    End Sub

    ' Saves the tab group list and any changes you have made.
    Sub Save
    End Sub

    ' Updates the TabGroups object to reflect any changes made through the Preferences user interface.
    Sub Update
    End Sub
End Class

'This object represents a folder tab in a tab group.
Class TabGroupTabEntry
    ' Returns the color, if any, assigned to this tab.
    Property Get color ' Return Type string
    End Property

    ' Returns the folder format of this tab.
    Property Get format ' Return Type object:Format 
    End Property

    ' Returns the link ID of this tab, if it is linked to another tab. Both tabs will have the same link ID but otherwise the value is meaningless. Use the TabGroup.Link and Unlink methods to change tab linkage.
    Property Get linkid ' Return Type int
    End Property

    ' If this tab is linked as a slave, returns the string ""slave"".
    Property Get linktype ' Return Type string
    End Property

    ' Returns the lock type of this tab. Valid values are ""on"", ""off"", ""changes"" and ""reuse"".
    Property Get locked ' Return Type string
    End Property

    ' Returns the name of this tab if one is assigned. Tabs that don't have specific names assigned will usually show the last component of the path as their name.
    Property Get name ' Return Type string
    End Property

    ' Returns the path that this tab will load when it's opened.
    Property Get path ' Return Type object:Path 
    End Property

    ' Returns a duplicate of this tab entry.
    Function Duplicate ' Return Type object:TabGroupTabEntry
    End Function
End Class

'This object represents a list of folder tabs in a tab group.
Class TabGroupTabList
    ' Returns a TabGroupTabEntry object representing the active (default) folder tab in this tab list.
    Property Get active ' Return Type object:TabGroupTabEntry
    End Property

    ' Adds a folder tab entry to this list. You can provide a
    Function AddTab(<object:TabGroupTabEntry> or <string:path>,<string:name>) ' Return Type object:TabGroupTabEntry
    End Function

    ' Deletes a folder tab entry from this list. You can provide a
    Sub DeleteTab(<object:TabGroupTabEntry> or <int:index>)
    End Sub

    ' Inserts a folder tab entry to this list. You can provide a
    Function InsertTabAt(<object:TabGroupTabEntry> or <string:path>,<string:name>,<int:index>) ' Return Type object:TabGroupTabEntry
    End Function

    ' Moves the specified tab entry to a new position, and optionally a new tab list. If the second parameter is a TabGroupTabList object then the tab entry will be moved to that list. The final parameter must be the index indicating the desired insertion position.
    Sub MoveTabTo(<object:TabGroupTabEntry>,<object:TabGroupTabList>,<int:index>)
    End Sub
End Class

'This object provides various statistics about a folder tab (the number of selected files, total number of items, etc). It is obtained from the Tab.stats and Tab.selstats properties.
Class TabStats
    ' Returns the width in pixels of the largest image in the folder.
    Property Get bigimage_h ' Return Type int
    End Property

    ' Returns the height in pixels of the largest image in the folder.
    Property Get bigimage_w ' Return Type int
    End Property

    ' Returns the total number of bytes in the folder as a FileSize object.
    Property Get bytes ' Return Type object:FileSize 
    End Property

    ' Returns True if the tab is currently in Checkbox Mode.
    Property Get checkbox_mode ' Return Type bool
    End Property

    ' Returns the total number of bytes in checked items as a FileSize object.
    Property Get checkedbytes ' Return Type object:FileSize 
    End Property

    ' Returns the total number of bytes in checked folders as a FileSize object.
    Property Get checkeddirbytes ' Return Type object:FileSize 
    End Property

    ' Returns the total number of checked folders.
    Property Get checkeddirs ' Return Type int
    End Property

    ' Returns the total number of bytes in checked files as a FileSize object.
    Property Get checkedfilebytes ' Return Type object:FileSize 
    End Property

    ' Returns the total number of checked files.
    Property Get checkedfiles ' Return Type int
    End Property

    ' Returns the total number of checked items.
    Property Get checkeditems ' Return Type int
    End Property

    ' Returns the total length in seconds of all checked music files.
    Property Get checkedmusiclength ' Return Type int
    End Property

    ' Returns the total number of bytes in all folders as a FileSize object.
    Property Get dirbytes ' Return Type object:FileSize 
    End Property

    ' Returns the total number of folders.
    Property Get dirs ' Return Type int
    End Property

    ' Returns the total number of bytes in all files as a FileSize object.
    Property Get filebytes ' Return Type object:FileSize 
    End Property

    ' Returns the latest (most recent) file date in the folder.
    Property Get filedate_max ' Return Type date
    End Property

    ' Returns the earliest (oldest) file date in the folder.
    Property Get filedate_min ' Return Type date
    End Property

    ' Returns the total number of files.
    Property Get files ' Return Type int
    End Property

    ' Returns the total number of items.
    Property Get items ' Return Type int
    End Property

    ' Returns the size of the largest file in the folder as a FileSize object.
    Property Get largestfile ' Return Type object:FileSize 
    End Property

    ' Returns the total length in seconds of all music files.
    Property Get musiclength ' Return Type int
    End Property

    ' Returns the total number of bytes in all selected items as a FileSize object.
    Property Get selbytes ' Return Type object:FileSize 
    End Property

    ' Returns the total number of bytes in all selected folders as a FileSize object.
    Property Get seldirbytes ' Return Type object:FileSize 
    End Property

    ' Returns the number of selected folders.
    Property Get seldirs ' Return Type int
    End Property

    ' Returns the total number of bytes in all selected files as a FileSize object.
    Property Get selfilebytes ' Return Type object:FileSize 
    End Property

    ' Returns the number of selected files.
    Property Get selfiles ' Return Type int
    End Property

    ' Returns the number of selected items.
    Property Get selitems ' Return Type int
    End Property

    ' Returns the total length in seconds of all selected music files.
    Property Get selmusiclength ' Return Type int
    End Property

    ' The first time a script accesses a particular TabStats object, a snapshot is taken of the tab state. If the script then makes changes to that tab (e.g. it selects files, creates a new folder, etc), these changes will not be reflected by the object. To re-synchronize the object with the tab, call the TabStats.Update method.
    Sub Update
    End Sub
End Class

'This object represents a toolbar. It is obtained with the DOpus.toolbars and Lister.toolbars properties.
Class Toolbar ' Default returns the name of the toolbar as string.
    ' Returns True if this is a default (factory-provided) toolbar, or False if it was user-created.
    Property Get deftoolbar ' Return Type bool
    End Property

    ' Returns a collection of Lister objects representing any and all Listers this toolbar is currently open in.
    Property Get listers ' Return Type collection:Lister 
    End Property

    ' Returns a collection of Dock objects representing any currently floating instances of this toolbar.
    Property Get docks ' Return Type collection:Dock 
    End Property

    ' Returns a string indicating the group (position) of a particular instance of this toolbar. The returned string will be one of top, bottom, left, right, center, fdright, fdbottom, tree.
    Property Get group ' Return Type string
    End Property

    ' Returns the line number within the toolbar's group that it resides on. For example, the first toolbar at the top of the Lister would have a line of 0.
    Property Get line ' Return Type int
    End Property

    ' Returns the pixel position from the left/top of the toolbar's line. If there are two or more toolbars with the same line number, the pos value determines the order they appear in.
    Property Get pos ' Return Type int
    End Property
End Class

'The Toolbars object lets you enumerate all the defined toolbars in your Directory Opus configuration (whether currently turned on or not).
Class Toolbars ' Default returns a collection of Toolbar objects that you can enumerate.
    ' Returns the name(s) of the currently selected File Display Toolbar(s).
    Property Get fdb ' Return Type string
    End Property

    ' Returns the name of the currently selected Viewer Toolbar.
    Property Get viewer ' Return Type string
    End Property
End Class

'Similar to a StringSet but can store elements of any type rather than just strings.
Class UnorderedSet
    ' Returns the number of elements the UnorderedSet currently holds.
    Property Get count ' Return Type int
    End Property

    ' Returns True if the UnorderedSet is empty, False if not.
    Property Get empty ' Return Type bool
    End Property

    ' A synonym for count.
    Property Get length ' Return Type int
    End Property

    ' A synonym for count.
    Property Get size ' Return Type int
    End Property

    ' Copies the contents of another UnorderedSet to this one. You can also pass an array or Vector object.
    Sub assign(<UnorderedSet:from>)
    End Sub

    ' Clears the contents of the UnorderedSet.
    Sub clear
    End Sub

    ' Erases the element if it exists in the set.
    Sub erase(variant)
    End Sub

    ' Returns True if the specified element exists in the set.
    Function exists(variant) ' Return Type bool
    End Function

    ' Inserts the element into the set if it doesn't already exist. Returns True if successful.
    Function insert(variant) ' Return Type bool
    End Function

    ' Merges the contents of another UnorderedSet with this one.
    Sub merge(<UnorderedSet:from>)
    End Sub
End Class

'This object represents a variable. Toolbar buttons, hotkeys and scripts can read and store variables, and variables can be saved from one session of Opus to another. The Var object is obtained from the Vars collection.
Class Var ' Default Return Type variant
    ' Returns the name of the variable. You cannot change the name of a variable once it has been assigned - instead, delete the variable from its collection and add a new one.
    Property Get name ' Return Type string
    End Property

    ' Returns True if the variable is persistent (saved) or False if not. You can set this property to change the persistence state.
    Property Get persist ' Return Type bool
    End Property

    ' Returns the value of the variable. You can set this property to change the value of the variable.
    Property Get value ' Return Type variant
    End Property

    ' Deletes this variable from its parent collection.
    Sub Delete
    End Sub

End Class

'This object represents a collection of variables. Depending on the variables' scope it can be obtained from the DOpus.vars, Lister.vars, Tab.vars, Command.vars or Script.vars properties.
Class Vars ' Default Return Type collection:Var 
    ' Deletes the named variable from the collection. You can also specify a wildcard pattern to delete multiple variables (or * for all).
    Sub Delete(<string:name>)
    End Sub

    ' Returns True if the named variable exists in the collection, or False if it doesn't exist.
    Function Exists(<string:name>) ' Return Type bool
    End Function

    ' Returns the value of the named variable. You can use this method as an alternative to indexing the collection.
    Function Get(<string:name>) ' Return Type variant
    End Function

    ' Sets the named value to the specified value. You can use this method as an alternative to indexing the collection.
    Sub Set(<string:name>,<variant:value>)
    End Sub
End Class

'This object is similar to an array - it can store an unlimited number of elements of any type. Several properties and methods in the Opus scripting interface use Vectors, and you can use them interchangeably with arrays in most cases. The Vector is provided because some scripting languages only offer incomplete or incompatible arrays - using Vectors means the object can be used consistently across any ActiveX scripting language. A Vector is created by the DOpusFactory.Vector method.
Class Vector
    ' Returns the capacity of the Vector (the number of elements it can hold without having to reallocate memory). This is not the same as the number of elements it currently holds, which can be 0 even if the capacity is something larger.
    Property Get capacity ' Return Type int
    End Property

    ' Returns the number of elements the Vector currently holds.
    Property Get count ' Return Type int
    End Property

    ' Returns True if the Vector is empty, False if not.
    Property Get empty ' Return Type bool
    End Property

    ' A synonym for count.
    Property Get length ' Return Type int
    End Property

    ' A synonym for count.
    Property Get size ' Return Type int
    End Property

    ' Copies the values of another Vector to the end of this one, preserving the existing values as well. If start and end are not provided, the entire Vector is appended - otherwise, only the specified elements are appended.
    Sub append(<Vector:from>,<int:start>,<int:end>)
    End Sub

   ' Copies the value of another Vector to this one. If start and end are not provided, the entire Vector is copied - otherwise, only the specified elements are copied.
    Sub assign(<Vector:from>,<int:start>,<int:end>)
    End Sub

    ' Returns the last element in the Vector.
    Function back ' Return Type variant 
    End Function

    ' Clears the contents of the Vector.
    Sub clear
    End Sub

    ' Erases the element at the specified index.
    Sub erase(<int:index>)
    End Sub

    ' Exchanges the positions of the two specified elements.
    Sub exchange (<int:index1>,<int:index2>)
    End Sub

    ' Returns the first element in the Vector.
    Function front ' Return Type variant
    End Function

    ' Inserts the provided value at the specified position.
    Sub insert(<int:index>,<variant:value>)
    End Sub

    ' Removes the last element of the Vector.
    Sub pop_back
    End Sub

    ' Adds the provided value to the end of the Vector.
    Sub push_back(<variant:value>)
    End Sub

    ' Reserves space in the Vector for the specified number of elements (increases its capacity, although the count of elements remains unchanged).
    Sub reserve(<int:capacity>)
    End Sub

    ' Resizes the Vector to the specified number of elements. Any existing elements past the new size of the Vector will be erased.
    Sub resize(<int:size>)
    End Sub

    ' Reduces the capacity of the Vector to the number of elements it currently holds.
    Sub shrink_to_fit
    End Sub

    ' Sorts the contents of the Vector. Strings and numbers are sorted alphabetically and numerically - other elements are grouped by type but not specifically sorted in any particular order.
    Sub sort
    End Sub

    ' Removes all but one of any duplicate elements from the Vector. The number of elements removed is returned.
    Function unique ' Return Type int
    End Function
End Class

'This object represents information about the current Opus version. It is obtained from the DOpus.Version property.
Class Version ' Default return full version string (as shown in the About dialog).
    ' The current build number.
    Property Get build ' Return Type int
    End Property

    ' The current module version (the version of dopus.exe itself). You can also enumerate or index this as a collection:int to retrieve the individual four digits of the module version.
    Property Get module ' Return Type string
    End Property

    ' The current product version (the release version of Directory Opus as a whole). You can also enumerate or index this as a collection:int to retrieve the individual four digits of the product version.
    Property Get product ' Return Type string
    End Property

    ' Returns a WinVer object which provides information about the current version of Windows.
    Property Get winver ' Return Type object:WinVer 
    End Property

    ' Returns True if the current version of Opus is the specified version or greater. You can specify the major version only (e.g. ""11""), a major and minor version (e.g. ""11.3"") or a specific beta version (e.g. ""11.3.1"").
    Function AtLeast(<string:version>) ' Return Type bool
    End Function
End Class

'This object provides metadata properties relating to movie files. It is obtained from the Metadata object.
Class VideoMeta
End Class

'The Viewer object represents a standalone image viewer.
Class Viewer
    ' Returns the bottom coordinate of the viewer window.
    Property Get bottom ' Return Type int
    End Property

    ' Returns an Item object representing the currently displayed image.
    Property Get current ' Return Type object:Item 
    End Property

    ' Returns a collection of Item objects representing the images in the viewer's list.
    Property Get files ' Return Type collection:Item 
    End Property

    ' Returns True if the viewer is currently the foreground (active) window in the system.
    Property Get foreground ' Return Type bool
    End Property

    ' Returns the index of the currently viewed image within the viewer's list of files.
    Property Get index ' Return Type int
    End Property

    ' Returns True if the viewer is the most recently active viewer.
    Property Get lastactive ' Return Type bool
    End Property

    ' Returns the left coordinate of the viewer window.
    Property Get left ' Return Type int
    End Property

    ' Returns a Tab object representing the tab that launched the viewer (if there was one, and if it still exists).
    Property Get parenttab ' Return Type object:Tab 
    End Property

    ' Returns the right coordinate of the viewer window.
    Property Get right ' Return Type int
    End Property

    ' Returns or sets the title bar string for the viewer window.
    Property Get title ' Return Type string
    End Property

    ' Returns the top coordinate of the viewer window.
    Property Get top ' Return Type int
    End Property

    ' Adds the specified file to the viewer's current list of files. You can either pass a string or a Path object to indicate the file to add to the list. By default the file will be added to the end of the list, unless you specify a 0-based index as the second argument.
    Sub AddFile(<string:filepath>,<int:index>)
    End Sub

    ' Runs a command in the context of this viewer window. You can either pass a string or a Command object.
    Sub Command(<string:command> or <Command:command>)
    End Sub

    ' Removes the specified file from the viewer's current list of files. You can either pass the 0-based index of the file to remove, or the filepath (either as a string or a Path object).
    Sub RemoveFile(<int:index> or <string:filepath>)
    End Sub

    ' Used to change how the viewer window is grouped with other Opus windows on the taskbar. Specify
    Function SetTaskbarGroup ' Return Type bool
    End Function
End Class

'The Viewers object is a collection of all currently open standalone image viewers. It can be obtained via the DOpus.viewers property
Class Viewers ' Default return Type collection:Viewer
    ' Returns a Viewer object representing the most recently active viewer window.
    Property Get lastactive ' Return Type object:Viewer 
    End Property
End Class

'This object allows a script to access the in-built pattern matching functions in Opus. It is obtained from the FSUtil.NewWild method.
Class Wild ' Default Returns the current pattern in the Wild object, Type string.
    ' Escapes all wildcard characters in the input string and returns the result. For example, ""the * 'dog' said *"" would be conterted to ""the '* ''dog'' said '*"".
    Function EscapeString(<string:input>,<string:type>) ' Return Type string
    End Function

    ' Compares the specified string against the previously-parsed pattern, and returns True if it matches.
    Function Match(<string:test>) ' Return Type bool
    End Function

    ' Parses the supplied pattern.
    Function Parse(<string:pattern>,<string:flags>) ' Return Type bool
    End Function
End Class

'This object represents information about the current Windows version. It is obtained from the Version.winver property.
Class WinVer ' Default Return Full Windows version string.
    ' True if running on a Server edition of Windows.
    Property Get server ' Return Type bool
    End Property

    ' True if running on Windows XP.
    Property Get xp ' Return Type bool
    End Property

    ' True if running on Windows XP or better (this will always be true).
    Property Get xporbetter ' Return Type bool
    End Property

    ' True if running on Windows Vista.
    Property Get vista ' Return Type bool
    End Property

    ' True if running on Windows Vista or better (later).
    Property Get vistaorbetter ' Return Type bool
    End Property

    ' True if running on Windows 7.
    Property Get win7 ' Return Type bool
    End Property

    ' True if running on Windows 7 or better.
    Property Get win7orbetter ' Return Type bool
    End Property

    ' True if running on Windows 8.
    Property Get win8 ' Return Type bool
    End Property

    ' True if running on Windows 8 or better.
    Property Get win8orbetter ' Return Type bool
    End Property

    ' True if running on Windows 8.1.
    Property Get win81 ' Return Type bool
    End Property

    ' True if running on Windows 8.1 or better.
    Property Get win81orbetter ' Return Type bool
    End Property

    ' True if running on Windows 10.
    Property Get win10 ' Return Type bool
    End Property

    ' True if running on Windows 10 or better.
    Property Get win10orbetter ' Return Type bool
    End Property
End Class



'This object is provided to the OnAboutScript method, which is called when the user clicks the About button for a script in the Toolbars / Scripts Preferences page.
Class AboutData
    ' This is a handle to the parent window that the script should use if displaying a dialog via the Dialog object. Even though this is not a Lister or Tab, it can still be assigned to the Dialog.window property to set the parent window of the dialog.
    Property Get window ' Return Type int
    End Property
End Class

'This object is provided to the OnActivateLister method, which is called whenever a Lister window is activated or deactivated.
Class ActivateListerData
    ' Returns True if this Lister is activating, False if deactivating. Note that if the activation moves from one Lister straight to another the script will be called twice.
    Property Get active ' Return Type bool
    End Property

    ' Returns a Lister object representing the Lister that is closing.
    Property Get lister ' Return Type object:Lister 
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property
End Class

'This object is provided to the OnActivateTab method, which is called whenever a tab is activated.
Class ActivateTabData
    ' Returns a Tab object representing the tab that has become active.
    Property Get newtab ' Return Type object:Tab 
    End Property

    ' Returns a Tab object representing the tab that has gone inactive.
    Property Get oldtab ' Return Type object:Tab 
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property
End Class

'This object is provided to the OnAddCommands method, which allows a script to add new internal commands.
Class AddCmdData
    ' Adds a new internal command to Opus. The returned ScriptCommand object must be properly initialized. A script add-in can add as many internal commands as it likes to the Opus internal command set.
    Function AddCommand ' Return Type object:ScriptCommand 
    End Function
End Class

'This object is provided to the OnAddColumns method, which allows a script to add new information columns.
Class AddColData
    ' Adds a new information column to Opus. The returned ScriptColumn object must be properly initialized. A script add-in can add as many columns as it likes, and these will be available in file displays, infotips and the Advanced Find function.
    Function AddColumn ' Return Type object:ScriptColumn 
    End Function
End Class

'This object is provided to the OnAfterFolderChange method, which is called after a new folder has been read.
Class AfterFolderChangeData
    ' Returns a string indicating the action that triggered the folder read. The string will be one of the following: normal, refresh, refreshsub, parent, root, back, forward, dblclk.
    Property Get action ' Return Type string
    End Property

    ' If the read failed, this will return a Path object representing the path that Opus tried to read.
    Property Get path ' Return Type object:Path
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns True if the folder was read successfully, or False on failure.
    Property Get result ' Return Type bool
    End Property

    ' Returns a Tab object representing the tab that read the folder.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnBeforeFolderChange method, which is called before a new folder is read.
Class BeforeFolderChangeData
    ' Returns a string indicating the action that triggered the folder read. The string will be one of the following: normal, refresh, refreshsub, parent, root, back, forward, dblclk.
    Property Get action ' Return Type string
    End Property

    ' Returns True if this is the first path to be read into this tab (i.e. previously the tab was empty).
    Property Get initial ' Return Type bool
    End Property

    ' Returns a Path object representing the new path that is to be read.
    Property Get path ' Return Type object:Path 
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the tab that is changing folder.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnClick method, which is called whenever a script function is invoked (e.g. when the button is clicked or hotkey pressed).
Class ClickData
    ' Returns a Func object relating to this function. This provides access to information about the function's environment - (source and destination tabs, qualifier keys, etc).
    Property Get func ' Return Type object:Func 
    End Property

End Class

'This object is provided to the OnCloseLister method, which is called before a Lister closes.
Class CloseListerData
    ' Returns a Lister object representing the Lister that is closing.
    Property Get lister ' Return Type object:Lister 
    End Property

    ' Set this to True to prevent the closing Lister from being saved as the new default Lister.
    Property Get prevent_save ' Return Type bool
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns True if the Lister is closing because Opus is shutting down.
    Property Get shutdown ' Return Type bool
    End Property
End Class

'This object is provided to the OnCloseTab method, which is called before a tab closes.
Class CloseTabData
    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the tab that is closing.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnScriptConfigChange method, which notifies a script when the user edits the script's configuration.
Class ConfigChangeData
    ' Returns a Vector containing the names of the configuration items that were modified.
    Property Get changed ' Return Type Vector:string
    End Property
End Class

'This object is provided to the OnDisplayModeChange method, which is called when the display mode changes in a tab.
Class DisplayModeChangeData
    ' Returns a string indicating the new display mode. Will be one of largeicons, smallicons, list, details, power, thumbnails or tiles.
    Property Get mode ' Return Type string
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the tab the display mode changed in.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnDoubleClick method, which is called when a file or folder is double-clicked.
Class DoubleClickData
    ' Set this property to False to prevent the OnDoubleClick event being called for any further files during this operation (this is only effective if more than one file was double-clicked). Any remaining files will be opened according to their default handlers.
    Property Get Call ' Return Type bool
    End Property

    ' Set this property to False to abort double-click processing altogether on any further files during this operation (this is only effective if more than one file was double-clicked).
    Property Get cont ' Return Type bool
    End Property

    ' Returns True if your OnDoubleClick event is being called with only a path (via the path property) and not a full Item object. This will occur if you set the ScriptInitData.early_dblclk property to True when initialising your script.
    Property Get early ' Return Type bool
    End Property

    ' Returns True if the item double-clicked is a directory, False if it's a file.
    Property Get is_dir ' Return Type bool
    End Property

    ' Returns a Item object representing the item that was double-clicked. This property is only present if the early property is False.
    Property Get item ' Return Type object:Item 
    End Property

    ' Returns a string that indicates the mouse button that launched the double-click. The string can be one of the following: left, middle, none.
    Property Get mouse ' Return Type string
    End Property

    ' This is set to True if multiple files were double-clicked.
    Property Get multiple ' Return Type bool
    End Property

    ' Returns a Path object providing the full pathname of the item that was double-clicked.
    Property Get path ' Return Type object:Path
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' When the early property is True, set skipfull to True to prevent your OnDoubleClick event from being called a second time.
    Property Get skipfull ' Return Type bool
    End Property

    ' Returns a Tab object representing the tab that the item was double-clicked in.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnFileOperationComplete method, which lets you receive notification when certain file operations complete.
Class FileOperationCompleteData
    ' Returns a string that indicates the type of file operation. Currently the only supported value is ""rename"".
    Property Get action ' Return Type string
    End Property

    ' Returns a string that provides the entire command line that launched this operation.
    Property Get cmdline ' Return Type string
    End Property

    ' When the query property is False this provides further information about the operation that completed.
    Property Get data ' Return Type variant
    End Property

    ' Returns a Path object representing the destination path of the operation. 
    Property Get dest ' Return Type object:Path
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the operation was initiated.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns True the first time the OnFileOperationComplete event is called. You should examine the action and other properties and return True if you decide you want notification about this operation. This will be False when you are called the second time, when the operation is complete.
    Property Get query ' Return Type bool
    End Property

    ' Returns a Path object representing the source path of the operation.
    Property Get source ' Return Type object:Path
    End Property

    ' Returns a Tab object representing the source folder tab.
    Property Get tab ' Return Type object:Tab
    End Property
End Class

'This object is provided to the OnFlatViewChange method, which is called when the Flat View mode changes in a tab.
Class FlatViewChangeData
    ' Returns a string indicating the new Flat View mode. Will be one of off, grouped, mixed or mixednofolders.
    Property Get mode ' Return Type string
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the tab the Flat View mode changed in.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnGetCopyQueueName event, which is called whenever a copy operation begins that uses automatically-managed copy queues.
Class GetCopyQueueNameData
    ' Returns a Path object representing the destination path of the copy operation.
    Property Get dest ' Return Type object:Path 
    End Property

    ' Returns a Tab object representing the destination folder tab.
    Property Get desttab ' Return Type object:Tab
    End Property

    ' Returns a binary string indicating the physical drive indices that the destination path is located on (if any). For example, 00100000000000000000000000 indicates that drive C: is the destination drive.
    Property Get dest_drives ' Return Type string
    End Property

    ' Returns True if the operation is a move instead of a copy. 
    Property Get move ' Return Type bool
    End Property

    ' Returns the default queue name for this operation.
    Property Get name ' Return Type string
    End Property

    ' Returns a Path object representing the source path of the copy operation.
    Property Get source ' Return Type object:Path  
    End Property

    ' Returns a Tab object representing the source folder tab.
    Property Get sourcetab ' Return Type object:Tab
    End Property

    ' Returns a binary string indicating the physical drive indices that the source path is located on (if any). For example, 00001000000000000000000000 indicates that drive E: is the source drive.
    Property Get source_drives ' Return Type string
    End Property
End Class

'This object is provided to the OnGetCustomFields event, which lets a rename script add its own fields to the Rename dialog.
Class GetCustomFieldData
    ' Returns a CustomFieldData object, that the script can use to add custom fields to the Rename dialog. Each property added to the object in this method will be create a new field in the dialog, allowing the user to supply additional information to your rename script.
    Property Get fields ' Return Type object:CustomFieldData 
    End Property

    ' This lets you assign labels to your script's custom fields, that are shown to the user in the Rename dialog. To do this, set this property to a Map created via the DOpusFactory.Map method, filled with name/label string pairs.
    Property Get field_labels ' Return Type object:Map 
    End Property

    ' This lets you assign ""cue banners"" to any edit fields created by your script. A cue banner is displayed inside an empty edit field to prompt the user what sort of data the field expects. To use this, set this property to a Map created via the DOpusFactory.Map method, filled with name/banner string pairs.
    Property Get field_tips ' Return Type object:Map 
    End Property

    ' You can use this field to specify which control gets the input focus by default when your fields appear for the first time. Set it to the name of the desired control. You can also specify !oldname or !newname to assign focus to the standard old and new name fields.
    Property Get focus ' Return Type string
    End Property
End Class

'This object is provided to the OnGetHelpContent event, which lets a script add its own content to the Opus F1 help.
Class GetHelpContentData
    ' Adds a PNG or JPG image and makes it available for your help pages. You can use any name you like for your images, although they must have either a .png or a .jpg suffix. Your help content can then refer to images by name, e.g. if you add an image and call it myimage.jpg, your html content could show it using:
    Sub AddHelpImage(<string:name>,<Blob:image>) ' Return Type none
    End Sub
    ' Adds a page of help content for your script to the F1 help file. You can call this method as many times as you like. If you add more than one page of help the first page will become the topic header and all subsequent pages will appear underneath it in the index.
    Sub AddHelpPage(<string:name>,<string:title>,<string:body>) ' Return Type none
    End Sub
End Class

'This object is provided to the OnGetNewName method, which is one of the supported methods a rename script can provide.
Class GetNewNameData
    ' Returns a CustomFieldData object which provides the values of any custom fields your script added to the Rename dialog.
    Property Get custom ' Return Type object:CustomFieldData 
    End Property

    ' Returns an Item object representing the file or folder being renamed.
    Property Get item ' Return Type object:Item
    End Property

    ' Returns the proposed new name of the item. This will be the result of the application of any selected standard options in the rename dialog (numbering, capitalization, etc).
    Property Get newname ' Return Type string
    End Property

    ' Returns the file extension of the proposed new name. Does not take multi-part extensions into account (e.g. will return "".rar"" rather than "".part1.rar"").
    Property Get newname_ext ' Return Type string
    End Property

    ' Returns the file extension of the proposed new name, taking multi-part extensions into account (e.g. will return "".part1.rar"" rather than "".rar"").
    Property Get newname_ext_m ' Return Type string
    End Property

    ' Returns the contents of the New Name field (that is, not the calculated new name after all the options have been applied, but the actual text contents of the field as entered by the user).
    Property Get newname_field ' Return Type string
    End Property

    ' Returns the file stem of the proposed new name. Does not take multi-part extensions into account (e.g. will return ""catpictures.part1"" rather than ""catpictures"").
    Property Get newname_stem ' Return Type string
    End Property

    ' Returns the file stem of the proposed new name, taking multi-part extensions into account (e.g. will return ""catpictures"" rather than ""catpictures.part1"").
    Property Get newname_stem_m  ' Return Type string
    End Property

    ' Returns the ""old name"" pattern as entered by the user in the rename dialog.
    Property Get oldname_field ' Return Type string
    End Property
End Class

'This object is provided to the OnListerResize event, which is called whenever a Lister window is resized.
Class ListerResizeData
    ' Returns a string indicating the resize action that occurred. This will be one of the following strings: resize, minimize, maximize, restore.
    Property Get action ' Return Type string
    End Property

    ' Returns the new width of the Lister in pixels.
    Property Get width ' Return Type int
    End Property

    ' Returns the new height of the Lister in pixels.
    Property Get height ' Return Type int
    End Property

    ' Returns a Lister object representing the Lister that was resized.
    Property Get lister ' Return Type object:Lister 
    End Property
End Class

'This object is provided to the OnListerUIChange method, which is called when various user interface elements (tree, viewer, etc) are open or closed in a Lister.
Class ListerUIChangeData
    ' Returns a string indicating which UI elements changed. This will contain one or more of the following strings: dual, tree, metapane, viewer, utility, duallayout, metapanelayout, viewerlayout, toolbars, toolbarset, toolbarsauto, minmax.
    Property Get change ' Return Type string
    End Property

    ' Returns a Lister object representing the Lister that is changing.
    Property Get lister ' Return Type object:Lister 
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property
End Class

'This object is provided to the OnOpenLister method, which is called when a new Lister is opened.
Class OpenListerData
    ' Initially this is set to False, indicating that the event has been called before any tabs have been created. If you return True from the OnOpenLister event, it will be called again and after will be set to True to indicate all tabs have been created.
    Property Get after ' Return Type bool
    End Property

    ' Returns a Lister object representing the newly opened Lister.
    Property Get lister ' Return Type object:Lister 
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property
End Class

'This object is provided to the OnOpenTab method, which is called when a new tab is opened.
Class OpenTabData
    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the newly opened tab.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnInit method, which is called once to initialize each script in the Script Addins folder.
Class ScriptInitData
    ' Returns a ScriptConfig object, that the script can use to initialize its default configuration. Properties added to the object in this method will be displayed to the user in Preferences, allowing them to change their value and thus configure the behavior of the script.
    Property Get config ' Return Type object:ScriptConfig 
    End Property

    ' This lets you assign descriptions for your script's configuration options that are shown to the user in the editor dialog. To do this, set this property to a Map created via the DOpusFactory.Map method, filled with name/description string pairs.
    Property Get config_desc ' Return Type object:Map 
    End Property

    ' This lets you organize your script's configuration options into groups when shown to the user in the editor dialog. The group names are arbitrary - configuration options with the same group name will appear grouped together. Set this property to a Map created via the DOpusFactory.Map method, filled with name/group string pairs.
    Property Get config_groups ' Return Type object:Map 
    End Property

    ' Lets the script specify a copyright message that is displayed to the user in Preferences.
    Property Get copyright ' Return Type string
    End Property

    ' Set this to True if the script should be enabled by default, or False if it should be disabled by default. The user can enable or disable scripts using Preferences - this simply controls the default state.
    Property Get default_enable ' Return Type bool
    End Property

    ' Lets the script specify a description message that is displayed to the user in Preferences.
    Property Get desc ' Return Type string
    End Property

    ' Set this to True if your script implements the OnDoubleClick event and (for performance reasons) you want to be called with only a path to the double-clicked item rather than a full Item object. See the OnDoubleClick event documentation for more details.
    Property Get early_dblclk ' Return Type bool
    End Property

    ' Returns the path and filename of this script.
    Property Get file ' Return Type string
    End Property

    ' Lets you specify an arbitrary group for this script. If scripts specify a group they will be displayed in that group in the list in Preferences.
    Property Get group ' Return Type string
    End Property

    ' Lets the script specify a string that will be prepended to any log output it performs. If not set the name of the script is used by default.
    Property Get log_prefix ' Return Type string
    End Property

    ' Specifies the minimum Opus version required. If the current version is less than the specified version the script will be disabled. You can specify the major version only (e.g. ""11""), a major and minor version (e.g. ""11.3"") or a specific beta version (e.g. ""11.3.1"" for 11.3 Beta 1).
    Property Get min_version ' Return Type string
    End Property

    ' Lets the script specify a display name for the script that is shown in Preferences.
    Property Get name ' Return Type string
    End Property

    ' The OnInit method is called in two different circumstances - once during Opus startup, and again if the script is installed or edited when Opus is already running. This property will return True if the OnInit method is being called during Opus startup, or False for any other time.
    Property Get startup ' Return Type bool
    End Property

    ' Lets you provide a URL where the user can go to find out more about your script (it's displayed to the user in Preferences).
    Property Get url ' Return Type string
    End Property

    ' Returns a Vars collection of user and script-defined variables that are local to this script. These variables are available to other methods in the script via the Script.vars property.
    Property Get vars ' Return Type object:Vars 
    End Property

    ' Lets the script specify a version number string that is displayed to the user in Preferences.
    Property Get version ' Return Type string
    End Property

    ' Adds a new information column to Opus. The returned ScriptColumn object must be properly initialized. A script add-in can add as many columns as it likes, and these will be available in file displays, infotips and the Advanced Find function.
    Function AddColumn ' Return Type object:ScriptColumn 
    End Function

    ' Adds a new internal command to Opus. The returned ScriptCommand object must be properly initialized. A script add-in can add as many internal commands as it likes to the Opus internal command set.
    Function AddCommand ' Return Type object:ScriptCommand 
    End Function
End Class

'This object is provided to the OnShutdown method, which is called before Opus shuts down.
Class ShutdownData
    ' Returns True if the Windows session is ending (that is, if Opus is shutting down because the system is shutting down), or False if it's just Opus that is quitting.
    Property Get endsession ' Return Type bool
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property
End Class

'This object is provided to the OnSourceDestChange method, which is called when the source or destination state of a tab changes.
Class SourceDestData
    ' Returns True if the tab is now the destination.
    Property Get dest ' Return Type bool
    End Property

    ' Returns True if the tab is now the source. If both source and dest return False it indicates that the tab is now ""off"".
    Property Get source ' Return Type bool
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the tab.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnStartup method, which is called when Opus starts up.
Class StartupData
End Class

'This object is provided to the OnStyleSelected method, which is called when a new style is chosen in a Lister.
Class StyleSelectedData
    ' Returns a Lister object representing the Lister that changing style.
    Property Get lister ' Return Type object:Lister 
    End Property

    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns the name of the newly selected style.
    Property Get style ' Return Type string
    End Property
End Class

'This object is provided to the OnTabClick event, which is called whenever a tab is clicked with a qualifier key held down.
Class TabClickData
    ' Returns a string indicating any qualifier keys that were held down by the user when the event was triggered.
    Property Get qualifiers ' Return Type string
    End Property

    ' Returns a Tab object representing the tab that was clicked.
    Property Get tab ' Return Type object:Tab 
    End Property
End Class

'This object is provided to the OnViewerEvent event, which is called whenever certain events occur in a standalone image viewer.
Class ViewerEventData
    ' Returns a string indicating the event that occurred. The events currently defined are:create,destroy,load,setfocus,killfocus,click,dblclk,mclick
    Property Get event ' Return Type string
    End Property

    ' For the load event, returns an Item object representing the newly loaded image.
    Property Get item ' Return Type object:Item 
    End Property

    ' Returns a Viewer object representing the viewer the event occurred in.
    Property Get viewer ' Return Type object:Viewer 
    End Property

    ' For the click events, returns the x coordinate within the viewer window that the click occurred.
    Property Get x ' Return Type int
    End Property

    ' For the click events, returns the y coordinate within the viewer window that the click occurred. 
    Property Get y ' Return Type int
    End Property

    ' For the click events, returns the width of the viewer window.
    Property Get w ' Return Type int
    End Property

    ' For the click events, returns the height of the viewer window.
    Property Get h ' Return Type int
    End Property
End Class

' Function OnAboutScript(AboutData)
' End Function

' Function OnActivateLister(ActivateListerData)
' End Function

' Function OnActivateTab(ActivateTabData)
' End Function

' Function OnAddCommands(AddCmdData)
' End Function

' Function OnAddColumns(AddColData)
' End Function

' Function OnAfterFolderChange(AfterFolderChangeData)
' End Function

' Function OnBeforeFolderChange(BeforeFolderChangeData)
' End Function

' Function OnClick(ClickData)
' End Function

' Function OnCloseLister(CloseListerData)
' End Function

' Function OnCloseTab(CloseTabData)
' End Function

' Function OnScriptConfigChange(ConfigChangeData)
' End Function

' Function OnDisplayModeChange(DisplayModeChangeData)
' End Function

' Function OnDoubleClick(DoubleClickData)
' End Function

' Function OnFileOperationComplete(FileOperationCompleteData)
' End Function

' Function OnFlatViewChange(FlatViewChangeData)
' End Function

' Function OnGetCopyQueueName(GetCopyQueueNameData)
' End Function

' Function OnGetCustomFields(GetCustomFieldData)
' End Function

' Function OnGetHelpContent(GetHelpContentData)
' End Function

' Function OnGetNewName(GetNewNameData)
' End Function

' Function OnListerResize(ListerResizeData)
' End Function

' Function OnListerUIChange(ListerUIChangeData)
' End Function

' Function OnOpenLister(OpenListerData)
' End Function

' Function OnOpenTab(OpenTabData)
' End Function

' Function OnInit(ScriptInitData)
' End Function

' Function OnShutdown(ShutdownData)
' End Function

' Function OnSourceDestChange(SourceDestData)
' End Function

' Function OnStartup(StartupData)
' End Function

' Function OnStyleSelected(StyleSelectedData)
' End Function

' Function OnTabClick(TabClickData)
' End Function

' Function OnViewerEvent(ViewerEventData)
' End Function

