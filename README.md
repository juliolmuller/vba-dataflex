
# Dataflex Module for Visual Basic for Application

- **Developed by:** [Julio L. Muller](https://www.jjsolutions.net/)
- **Released on:** Jul 15, 2018
- **Updated on:** Jan 31, 2019
- **Latest version:** 1.0.0
- **License:** MIT

## Installation

Imprort the `*.bas` file into your Visual Basic project by following the steps:

1. With the Excel workbook open. start the VBE window (`Alt + F11`);
2. In the menu, click on *File* > *Import File...* (`Ctrl + M`);
3. Through the file explorer, select the **DataflexPCModule.bas** file;
4. An item called *DataflexPCModule* will show up on your *Modules* list;
5. Enjoy!

Alternatively, copy and paste the plain text from the `*.txt` file into an existing module in your project.

## Content Summary

| Type         | Name                                                                | Return Type |
|:------------:|:--------------------------------------------------------------------|:-----------:|
| **Sub**      | [Sleep](#manage-scripting-speed)                                     | -           |
| **Function** | [IsDataflexOpen](#switch-to-dataflex)                               | *Boolean*   |
| **Function** | [RunTaskInDataflex](#send-keyboard-instructions-to-the-application) | *Boolean*   |
| **Sub**      | [ToggleNumLock](#toggle-numlock)                                    | -           |

## Resources Documentation

Obviously, these resources are tricky and are far from being good practives on programming. It is a turn arround to automate tasks inwithin DataflexPC, like downloading reports and running jobpacks.

The principle here is to use keyboard commands to execute tasks in Dataflex. Therefore, it works as a simple emulator of the human action through the keyboard.

**Something that you must be attempting to** is to be logged into Dataflex already. Dataflex does not support keyboard navigation to the *login* button.

### Manage Scripting Speed

The `Sleep()` routine requires the processor scheduler to wait some time before keep processing the Visual Basic script. It is not necessary to be used on the `RunTaskInDataflex()` function, since it already makes use of that internally, but it is a function available to you, just in case.

#### Structure

```vbnet
Sub Sleep(dwMilliseconds As Long)
```

- **dwMilliseconds** - Amount of milliseconds to freeze.

### Switch to Dataflex

Function `IsDataflexOpen()` is a required step before submitting the commands. This looks for the instance of the application open in the Operating System and puts its window in foreground, ensuring it is active and readt for keyboard commands. You can use it in a *If* statement to capture its boolean return and take actions incase of any errors.

#### Structure

```vbnet
Function IsDataflexOpen() As Boolean
```

- ***return*** - Returns whther the operation was successful or not.

### Send Keyboard Instructions to the Application

To effectively initialize the sequence of instructions inside Dataflex, you will pass an array of *strings* and *integers* to the function `RunTaskInDataflex()`. This basically will execute two differenc instructions depending on the variable type:

- *Integers* are passed as paremeter to the subroutine `Sleep()`, indicating you want to hold on the amount of milliseconds before running the next instruction. This is highly recommended to be used when there are transitions of screens or submition of requests.
- *Strings* are passed as parameter to the subroutine `SendKeys()`. To set the appropriate string in the array, you must know the keys ID accepted by *SendKeys*, so it is good to take a look at the [official documentation](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/sendkeys-statement).

Besides, it is recommended not to run any more instructions after download requests, since the waiting time is very variable.

#### Structure

```vbnet
Function RunTaskInDataflex(arrInstructions As Variant) As Boolean
```

- **arrInstructions** - Receives an *array* with the instructions to be run.
- - ***return*** - The function returns a boolean value. Thus, you can evaluate its result in an *If* statement and treat errors.

### Toggle NumLock

The function `ToggleNumLock()` is available to adjust a well known (and anoying) issue with the usage of the **SendKeys** command: turning off the *Num Lock*. You can use this resource at the end of your script to ensure the *Num Lock* key is typed once.

#### Structure

```vbnet
Sub ToggleNumLock()
```

## Example

Below, you can see an example on dowloading a *Audit Trail Report*.

```vbnet
'Declare variable as "Variant"
    Dim commands As Variant

'Load variable with comands, using "Array()" function
    commands = Array( _
        "%SD", 1500, _
        udtInfoDFX.Cycle, "{TAB}", _
        udtInfoDFX.Period, "{TAB}", _
        "ALL", "{TAB}", _
        udtInfoDFX.RU, "{TAB}", _
        "{TAB}", 200, _
        "{ENTER}", 2000, _
        "+{TAB}+{TAB}+{TAB}", _
        200, _
        "{ENTER}")

'Ensure Dataflex window is in foreground
    If (IsDataflexOpen()) Then

'Execute instructions in Dataflex
        If (RunTaskInDataflex(arrInstructions)) Then
            MsgBox "Success!", vbInformation
        End If
    End If

'Turns NUMLOCK back on
    Call ToggleNumLock
```

## Compatibility

The scripts were tested **ONLY** in MS Excel & Access 2013. Other MS Office applications were not tested.

Please, report any issues (or even success on running it in other applications or MS Office versions) through the commentary session.

## Other Contents

- [Send Emails with Outlook - Module for VBA](https://github.com/juliolmuller/VBA-Module-Outlook)
- [File Handler Module for VBA](https://github.com/juliolmuller/VBA-Module-TextFile)
