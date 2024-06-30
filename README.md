# Easy Material Weight Management for Siemens NX

an NX Study for Joinery Designers and others: Mastering Weight Management through Journals, Component Creation, Dimension Tool and more without any licences

![MyEWSetup](https://github.com/lumberjack1970/EasyWeightforNX/assets/164236127/9022dba2-48ba-4267-b342-0145c7edc4ca)

## How to use

Every journal can be easily customized to meet your preferences.    
I recommend using Notepad++ or Visual Studio Code. To design icons, Krita is an excellent choice. (any of these are free)    
https://notepad-plus-plus.org/    
https://krita.org/    

1. Open any journal for editing. Be brave, they are all in plain English with comments on the know-how.
2. At the beginning, you'll find a section labeled 'Configuration settings'.
3. These settings allow you to easily change the main features without diving into the code. (Only modify the code directly if you really know what you're doing and couldn't find what you need.)
   
Simply adjust the settings as needed:

- Use 'true' for yes/on
- Use 'false' for no/off
- For specific settings, you'll find an explanation on how to set them up

To create custom buttons, follow these simplified steps:

1. Open Customize Dialog: Right-click in the menu or toolbar area and select "Customize."
2. Add New Button: In the "Customize" dialog, go to the "Commands" tab, find "New Item" near the bottom of the list, and drag the "New User Command" to where you want it in the toolbar. A visual indicator shows where it will be placed.
3. Customize Button: Right-click the new button to edit its name, icon, and action. To link the button to a journal, select "Edit Action," change the type to "Journal File," and browse to your journal file. Select Visual Basic Files (.vb) Confirm your settings.
   
Now your custom button is ready to use, directly executing your journal with a single click.

## Weight Calculation in a Multi-body environment for solid bodies
I created several versions of the Material Journal and placed them on the ribbon with a corresponding icon.

1. **Material Journal**   

   - _EW_Material_12mm Plywood.vb_: Changes body color, layer and transparency, sets a density value, measures volume, calculates weight, and attributes it: EW_Material, EW_Material_Density and EW_Body_Weight.

   - _NX_Material_12mm Plywood.vb_: Changes body color, layer and transparency and sets an NX's built in material. You have to create your own Material Library **- physicalmateriallibrary.xml -** to use this. See below for further details.

2. **Face Material Journal** - _EW_Material_FACE_Inside.vb_:  
Alters the color of selected faces. Has priority over the main Material Journal. Used to distinguish the inside/outside of the body.

3. **Raw Body Journal** - _EW_Material_RAW BODY.vb and NX_Material_RAW BODY.vb_:  
By selecting the original body and the raw body, this calculates the weight difference and adds a new attribute: Raw_Body_Delta_Weight. It also moves the raw body to a predefined layer and makes it transparent. Useful to handle this on a custom level.

4. **Delete Attributes Journal** - _EW_Material_DELETE ATTRIBUTES.vb and NX_Material_NULLMATERIAL.vb_:   
Keeps the body unchanged but removes any weight-related attributes or sets a "NullMaterial" with zero weight respectively. Created, so these bodies won’t be included in the weight calculation on the drawing.

### Solid Body Material Filter Tool
Using this tool, you can control the visibility of specific solid bodies on your screen using the attributes assigned before. When creating components, it simplifies the process of organizing them. This tool automatically adjusts visibility based on the chosen materials. If no attribute is found, it hides them among the others. "Without Weight" option at the bottom displays all bodies that lack weight information. This allows you to double-check your work.
This journal needs setup like the others too. Have all your materials ready to fill up its database, so it will know what to look for. You will probably need to tailor the window size, which you can find at the end of the journal. I've left comments for you to know which numbers need to be changed.

### Component Creator 
This tool enables you to automatically create parts by requesting you a main component name. For example, "MyProject-01" creates: MyProject-01-101-Top Panel, MyProject-01-102-Shelf, etc. Select solid bodies to create components for. See below for further details.   

### Component Weight Transfer
In the Modeling environment/Main Assembly, this journal transfers weight information (weight attribute - EW_Body_Weight) from solid bodies to components. Summarizes all component weights to assign a Total Assembly Weight attribute to the Main Assembly, excluding weights of underlying components. To be used exclusively with the original - EW_Material_12mm Plywood. When you assign one of NX's built-in materials, this function occurs natively.

### Total Weight to Drawings
In the Drafting environment, sums all solid body weights for a Total Built-in Weight and adds Raw body differences for a Total Environmental Weight in the title block. Does not require Component Weight Transfer Journal.

-----

> [!IMPORTANT]
>  For EasyWeight (**EW_Material_12mm Plywood.vb**) users:
> - Weight is calculated during the Material Journal to a solid body.
> - Journals are not associative. Any geometry changes require Journal reapplication.
> - Component Creator updates all relevant EasyWeight information by default.

---

## ...And some additional journals to get you through the day

### Dimensions Tool
Automates dimensions - Lenght, Width and Material thickness in components for aligned and non-aligned solid bodies. See below for further details.           

### Layer Quick Switch Journal
Originally designed to act as a toggle switch, but it became an impossible mission after numerous failed attempts to make it work.   
As a result, two journals were created: one for 'on' and one for 'off'. It can handle multiple layers though.

```vbnet
Configuration Code Snippets:

layerNumbers As Integer() = {1, 70, 90}  ' Add or remove layer numbers as needed, separate with coma
layerState As NXOpen.Layer.State = NXOpen.Layer.State.Hidden  ' Can be set to Visible or Hidden
```

### Construction Objects Visibility Journal
This enhanced version of the previous Layer Quick Switch Journal allows you to easily toggle the visibility of various construction objects, including Sketches, Curves, Datums, Routing, Assembly Constraints, and Layers. It provides separate settings to manage which objects need to be visible and which need to be hidden.

### Drafting View Border On / Off Journal
Created, because it was a consistent struggle to locate the drawing border when I hovered my mouse.   
You can find this function under Drafting preferences / View / Workflow / Border.

# Detailed Information and Configuration Settings

### Material Journals

* **Which should you use** - _EW_Material_12mm Plywood.vb_ **or** _NX_Material_12mm Plywood.vb_**?**

It's important to understand that the Easyweight project originated from a straightforward concept: circumventing NX's limitations when using it without a material license, allowing for body modifications and the assignment of a material name. As the project evolved, I discovered that it was also possible to assign built-in materials, presenting you with two options. Ideally, I should deprecate the first one, but I've chosen to keep it because it represents the original concept — and I love its simple and elegant solution to such limitations. The choice is ultimately yours, but I encourage you to develop your own built-in material library and use the second - **associative** - option. Every subsequent journal is prepared to accept either one or has an alternative.
                           
### EW_Material_12mm Plywood.vb

  How to use:

1. Establish Your Color System in NX    
Navigate to **Menu → Preferences → Color Palette**.    
Hover your mouse over any color to see its ID number, which you'll need later.

2. Create a Journal for Each Material    
Save as and start editing. At the beginning of the code, where the Configuration Settings appear, enter your desired material details:

```vbnet
Configuration Code Snippets:

materialName As String = "12mm Plywood"
density As Double = 440 ' Kg/m3 or Pound/Cubic Foot
unitSystem As String = "kg"
bodyColor As Double = 111 ' Set the solid body color to ID: 111
bodyLayer As Double = 1 ' Set the solid body to layer 1
bodyTranslucency As Double = 0 ' Set the solid body transparency to 0
```

3. Customize Your Ribbon with Buttons as [described earlier](#how-to-use)    
Create custom icons using the colors you’ve chosen for better visual recognition.
                    
### NX_Material_12mm Plywood.vb

  How to use:

1. Create a New Material Library    
Search the internet for "How to Create a New UG NX Material Library". Familiarize yourself with the following and create your own library as you like:
   - Copy, create and modify materials with the "Assign Materials" function.
   - Export materials to a library and save it as an *.xml file.

2. Establish Your Color System in NX    
Navigate to **Menu → Preferences → Color Palette**.    
Hover your mouse over any color to see its ID number, which you'll need later.

3. Create a Journal for Each Material    
Save as and start editing. At the beginning of the code, where the Configuration Settings appear, fill in the data you've previously gathered:

```vbnet
Configuration Code Snippets:

materialName As String = "12mm Plywood"
materialLibraryPath As String = "C:\Your Folder\To your material library\physicalmateriallibrary_custom.xml"
bodyColor As Double = 111
bodyLayer As Double = 1
bodyTranslucency As Double = 0
```

4. Customize Your Ribbon with Buttons as [described earlier](#how-to-use)    
Create custom icons using the colors you’ve chosen for better visual recognition.

##
|  Available Journals      | _EW_Material_12mm Plywood.vb_ | _NX_Material_12mm Plywood.vb_ |
|:-------------------------|:-----------------------------:|:-----------------------------:|
| Face Material | Yes | Yes |
| Raw Body | EW_Material_RAW BODY.vb | NX_Material_RAW BODY.vb |
| Delete Attributes | EW_Material_DELETE ATTRIBUTES.vb | NX_Material_NULLMATERIAL.vb |
| Solid Body Material Filter | EW_Material Filter | NX_Material Filter |
| Component Creator | Yes - See Code to Setup | Yes - See Code to Setup |
| Component Weight Transfer | Yes | Not Applicable |
| Total Weight to Drawings | EW_Total Weight to Drawing | NX_Total Weight to Drawing |
| Dimensions Tool | Not Related | Not Related |

-----
### Captain Hook's Component Creator

**Under the Hood**    
- Local: The tool creates the first sequential component labeled "-101". It then generates the next available number. These components aren't saved; they are created for you to save if you are satisfied with the outcome. The journal searches through your library, session, and memory for the component name you have set, checking for available or missing numbers to avoid duplication. The names for the components will be created from the names of the solid bodies. If a solid body does not have anything assigned, the default name “Panel” will be used. You can change this later in the code. See the Guide folder for screenshots on how to rename your solid bodies.

- Teamcenter: The tool employs a two-round approach to ensure each component is sequentially numbered and accurately tracked in Teamcenter. You can use both the first and second rounds or only the second, depending on your requirements. See the code for configuration examples.

   - **Round One:** The journal sets and saves an initial ID number for a component. This is crucial for establishing a base number that Teamcenter can recognize and refer to in subsequent operations. This step essentially anchors the starting point for a sequence of numbers and is necessary to follow a sequential order.   
_Example:_
For a specific component labeled "X184-500-101-PanelOne", the first round would involve creating this part with an ID of "X184-500-101" and saving it in Teamcenter. This saved ID becomes the reference point for the next numbers to be generated.

   - **Round Two:** The journal uses a substitute number to automatically generate subsequent IDs based on the initial number set in the first round. This is where the wildcard substitution becomes a bit of a game. The exact wildcard format can vary; sometimes, just a "\*" (star) might work, while other system has a more specific pattern like "-10*" might be necessary to correctly sequence the numbers.   
_Example:_
Continuing from the previous example, after "X184-500-101" is saved, you might try "X184-500-10\*" for the next part. Teamcenter recognizes the base ("X184-500-10") and automatically appends the next available sequential number, creating "X184-500-102" for the next component, and so on. It's important to experiment with different wildcard formats to find the one that Teamcenter responds to in your specific setup.

**Features**
- Smart Sorting: Leverages EasyWeight or NX's built-in material attributes to organize selected solid bodies by material name and weight in descending order, or retains the order of your initial selection.
- Unit Support: Offers support for both metric (millimeters) and imperial (inches) units in material names for smart sorting.
- EasyWeight Integration: For EasyWeight users, the tool updates all weight information before component creation with automatic unit system recognition.
- Flag created components to avoid duplication: Only for EasyWeight users as the material journal adds an empty 'Component created' attribute to solid bodies. If you have a proper Material licence, and want to use this journal with this function, you have to assign the attribute to all your solids first. Select all Solid Body / Properties / New Attribute / Title: Component created, Data Type: String. A more efficient way would be to change your default model template to have this. See the code for a detailed comment about the issue.
- Configuration Settings with detailed descriptions at the beginning of the Journal: WaveLink options, controlling numbering gaps, local environment and Teamcenter option:

```vbnet
Configuration Code Snippets:

defaultAssemblyid As String = "MyProject-01"
wavelinkFeature As Boolean = True
smartsortingFeatureQST As String = "Maybe"
ssUnitmm As String = "mm"
ssUnitin As String = "in"
EasyWeightsortinglogic As Boolean = True
defaultSolidbodyname As String = "PANEL"
setComponentflag As Boolean = False
teamcenterIntegrationQST As String = "False"
fillTheGap As Boolean = True
```
-----
### Lord Voldemort's Dimensions Tool - Length, Width And Material Thickness

**Under the Hood**
- Component Analysis: The script iterates through all components in the main assembly, including subassemblies, calculating the dimensions of each component. It intelligently handles duplicated components by skipping them.
- Bounding Box Calculation: For the designated body, the script will attempt to determine its width, depth, and height. The body doesn't have to be aligned with the absolute coordinate system. The process involves generating a non-aligned minimum bounding box, selecting the first vertex on it, iterating through the edges that share a common point with this vertex, and then measuring these three edges. Initially tries to determine the material thickness using the pre-set values. If it fails to find a match, the smallest value will be assigned. The longest edge will then be designated as the 'Length' and the remaining edge in the group will be identified as the 'Width'.

**Solid Body Requirements:**
- If the body is on Layer 1, it will be processed.
- If on any layer other than Layer 1, it will be skipped.
- If multiple bodies are on Layer 1, the script will skip the component.

**Features**
- User Interaction: A form interface allows for manual adjustments to dimensions and to change grain direction.
- Part-Level Unit Recognition: Metric (Millimeters) and Imperial (Inches) units at the part level within assemblies, with automatic adaptation to the specified unit system for each part.
- Measurement Precision Configuration.
- Nearest Half Rounding for Millimeters.
- Trim Trailing Zeros: A configuration to trim trailing zeros from formatted measurements for a cleaner numerical data presentation.
- GUI-Based Modification Control: You can toggle the setting to on, off, or to prompt you with a question. When enabled, it allows for interactive input modifications. When disabled, the program runs automatically with the predefined settings.
- Material Thickness Adjustment: A customizable setting for applying predefined adjustments to material thickness measurements.
- Configuration Settings with detailed descriptions at the beginning of the Journal:

```vbnet
Configuration Code Snippets:

validThicknesses As Double() = {0.141, 0.203, 0.25, 0.453, 0.469, 0.709, 0.827, 6, 9, 12, 13, 15, 18, 19}
decimalFormat As String = "F1"
roundToNearestHalfBLN As Boolean = True
modificationsQST As String = "Maybe"
materialThicknessadjust As Boolean = True
trimZeros As Boolean = True
LengthAttribute As String = "DIM_LENGTH"
WidthAttribute As String = "DIM_WIDTH"
MaterialThicknessAttribute As String = "MAT_THICKNESS"
```

> [!TIP]
> - NX Window Responsiveness: If you run the tool with interactive input modifications, the NX window become unresponsive during operation. Before initiating the process, set your Model to a trimetric view and arrange your Information Window in a way that it doesn't obstruct the view of your model. You can access the window using Ctrl+Shift+S.
> - Suppressing Components: If certain components don't require measurement, suppress them before running the tool to streamline the process.
> - Exiting the Tool: In the absence of an exit button, stop the code via NX to suspend the process. An error message will appear as part of the normal operation.
> - Understanding Tool Functionality: If you're uncertain about how this tool functions, I recommend trying the following: Use the 'Bounding Body' function, change the selection to 'Solid Body', then in the settings, select 'Create Non-Aligned Minimum Body' and choose a panel with an irregular shape. This will allow you to observe the tool in action and gain a rough understanding of its capabilities.

<table width="100%">
<tr>
<td width="25%" style="text-align: center">Component Creator</td>
<td width="25%" style="text-align: center">Dimensions Tool</td>
</tr>
<tr>
<td width="25%" style="text-align: center"><img src="https://github.com/lumberjack1970/EasyWeightforNX/assets/164236127/2c3241b0-9773-48d6-8fa6-06fb0ee53497"></td>
<td width="25%" style="text-align: center"><img src="https://github.com/lumberjack1970/EasyWeightforNX/assets/164236127/4a817b51-a180-491f-b76e-772ec76851de"></td>
</tr>
</table>

-----

# Thanks

To [NXJournaling.com](https://www.nxjournaling.com/content/easy-material-weight-management-part-1) and [Eng-Tips.com](https://www.eng-tips.com/viewthread.cfm?qid=514995) [2](https://www.eng-tips.com/viewthread.cfm?qid=516555) [3](https://www.eng-tips.com/viewthread.cfm?qid=518413) for providing invaluable examples, insightful comments, and educational content.   
And to ChatGPT, for advising and debugging the codes alongside me as my constant companion throughout this journey.
