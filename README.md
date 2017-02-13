# PPT Section Indicator

PPT Section indicator is a PowerPoint VSTO add-in that adds section indicators on presentations, adapting them dynamically to highlight the current section or slide. The motivation behind this add-in is to automate the rather boring task of copying/pasting and reformatting objects in every slide of a presentation just to create a simple progress tracker.

![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/screenshot.png)

**Read this document in its entirety before using the add-in**. This will help you understand how it works and prevent further problems.

**Table of Contents**

1. [Preconditions](#preconditions)
2. [Modes of operation](#modes-of-operation)
3. [Warnings](#warnings-important)
4. [Usage instructions](#usage-instructions)
  1. [Step 1 - Formatting](#step-1-formatting)
  2. [Step 2 - Positioning](#step-2-positioning)
5. [Download](#download) 
6. [Troubleshooting](#troubleshooting)

## Preconditions

For this tool to work, your presentation must be arranged into sections, as shown below.
![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/sections.png)

I've tested this tool only in PowerPoint 2016 (version 1612, from Dec. 2016), which is the version I have in my computer. This add-in was targeted to PowerPoint 2013, which, according to [Microsoft's documentation](https://msdn.microsoft.com/en-us/library/bb772080.aspx?f=255&MSPPError=-2147217396), should run in Office versions back to 2007.

## Modes of operation
This add-in has two modes of operation: (a) it can display section names only, or (b) track each slide individually through markers. Suppose, for instance, that we have a presentation with the following structure:

| Section name | Slides |
|--------------|--------|
| Intro        | 1-3    |
| Section 1    | 4-6    |
| Section 2    | 7-9    |

Mode (a) would produce the following results, depending on the current section of the slide.
![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/mode_a.gif)

Mode (b), in turn, would place markers for each slide, as shown below.
![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/mode_b.gif)

## Warnings (Important!!)
Before anything, I will list here a few things that you should be aware of before using this tool.

* First of all, **backup your presentation** before running the tool. It’s quite certain that I didn’t find all the bugs in it, and you may end up losing some information. Use it on your own risk;
* This add-in uses the clipboard to process objects among slides. Any content that you were holding on the clipboard will be lost once you run it. Also, **don't use the clipboard while the tool is running**;
* As an application-level add-in that needs to hold state, this tool doesn’t work well with multiple open presentations. The state seems to propagate between open instances of PowerPoint, making thigs quite messy. To avoid any problems, **use this tool only with a single instance of PowerPoint running**;
* Since this tool creates objects in the presentation, PowerPoint requires it to name them (this naming process happens every time you add something to your presentation, but PowerPoint picks some name automatically). To avoid clashes, I’ve decided to name each element starting with “SectionIndicator”, as shown in the selection pane below. If, by any chance, this tool complains about naming collisions and you’ve never run it before on your presentation, you managed to include an object in your presentation and name it starting with this reserved string. **Do not run the cleanup function of the tool**, as it will erase this object thinking that the tool itself created it in a previous run. Instead, find it on the selection pane and rename it.

![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/selection_pane.png )

## Settings
This add-in has a few straightforward settings:

**Include slide markers**: This checkbox alternates between the two modes of operation, as explained before in this document;

**Include hyperlinks**: Sets whether to treat section textboxes and slide markers as hyperlinks, pointing to their respective slides. When this option is enabled, you can click on the elements to navigate slides, as shown below.

![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/with_hyperlinks.gif )

**Slides to edit**: Determines which slides to include when processing the presentation. This option is useful if there are slides that are not to be included in the marker count (e.g., title slides, appendices, etc.). For instance, if the first slide of the presentation were the title, you could include slides 2 through 9 only. This affects the tool in the following way:

* Slide 1 won’t be processed. In other words, the progress indicator won’t be added in the first slide;
* The “Intro” section will have only two slides, which will impact the number of markers created by the tool.

The syntax through which you specify slide ranges is very similar to the one used when printing a document. Contiguous ranges are denoted by a dash ("-"), and distinct slide ranges or single slides are separated by semi-colons (";"). Here are some examples of valid expressions:

* 2-9 ; 11
* 2-9 ; 11-20 ; 22
* 2 ; 3 ; 4 ; 10-15

## Usage instructions

Once you divide the presentation in sections and specify the settings, press the “Start” button to begin. This tool works in two steps, as explained below.

### Step 1: Formatting

In this step, PPT Section Indicator will place some elements in the first slide of the range you specified previously. When you press the "Start" button, the tool is supposed to select this first slide, but I’ve found out that in some conditions (for instance, when the section where this first slide is placed is collapsed), the slide is not selected. In this case, you should select it manually.

Before proceeding, it’s important to define some naming conventions I’ve adopted in this tool. Suppose we have a presentation with the same structure as presented in the Modes of Operation section. For a given slide:

* The active section is the section where the slide is located. For instance, for slide number 1, the active section is "Intro";
* The inactive sections are all the remaining sections. For slide number 1, the inactive sections are "Section 1" and "Section 2".

The image below illustrates these concepts, assuming the current slide is slide number 1.

<img src="https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/elements_legend.png" width="500"/>

In the first step, the goal is to define the formatting of each element presented in the image. The tool should insert objects as shown below. 

![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/format_sample.png )

The purpose of each object should be clear at this point. The textboxes represent active and inactive sections, and (from left to right) the markers represent the current slide, slides in the active section and slides in inactive sections. The tool will use the style applied to those elements as a reference when creating the actual elements in the final step.

Here, you should define formatting aspects such as font size, font color, slide marker shape, size and color. Don’t worry about positioning elements in this step.

### Step 2: Positioning

One you’re done formatting the elements, press the "Next" button. The base textboxes and markers from the first step will be replaced by the formatted elements, considering the sections you’ve defined in your presentation. Now you should place the elements to your liking, taking into account that slide markers are ordered from left to right, starting a new line if necessary. It's also possible to resize elements (textbox and marker width/height), but you should not change anything else related to formatting, as it will be lost when propagating the objects to all slides. If you’d like to make changes, cleanup the presentation and start over.

Once you’re done placing the elements, press the "Done" button. PPT Section Indicator will propagate the elements to all slides in the specified range.

## Download

Visit the [Releases](https://github.com/fabio-muramatsu/PPTSectionIndicator/releases) page to find the most recent version. The installer is packed in ZIP format. Before extracting, it is important to trust the file, since Windows automatically recognizes it as downloaded from the internet and may prevent it from installing. To do that, right-click the ZIP file and, on the General tab, there's a Security field claiming that "The file came from another computer" and so on. Simply click on Unblock and you're done (more info [here](http://www.howtogeek.com/70012/what-causes-the-file-downloaded-from-the-internet-warning-and-how-can-i-easily-remove-it/)). Now, extract the contents and run the setup file.

If you have installed a previous version, you will have to remove it before reinstalling.

## Troubleshooting

#### When running the installer, I get the following message: "Customized functionality in this application will not work because the certificate used to sign the deployment manifest for SolutionName is not trusted. Contact your administrator for further assistance"

This happens because Office has a rather strict policy when installing solutions coming from the web (more info [here](https://msdn.microsoft.com/en-us/library/bb772087.aspx)). Read the [Download](#download) section for instructions on how to trust the installer.

#### The add-in misbehaves when more than one instance of PowerPoint is open

This is a known limitation of the tool. As a VSTO add-in, its state seems to be shared between instances of PowerPoint, and not isolated to each document. Because of that, I recommend that you only run one instance of PowerPoint when using this tool.

#### When pressing the "Start" button, PPT Section Indicator shows a message asking me to cleanup elements. What should I do?

This happens when the add-in found elements in your presentation whose names will clash with the ones used by the tool. As noted in the [Warnings](#warnings-important) section, I've decided to name each element starting with "SectionIndicator", so if PPT Section Indicator finds any element starting with this reserved string, it will show this warning. If you’ve run this tool before in your presentation and some elements created by it are still present, it is ok to clean them up and proceed. However, if you’ve never run the tool before, this means that some element is named starting with "SectionIndicator". Find this element using the selection pane and rename it.

#### When pressing the "Start button", I can't see the elements described in [Step 1](#step-1-formatting)

The add-in creates the elements in the first slide of the range you've specified. For instance, if the range is 2-9, the elements will be inserted in slide 2. This slide number is also shown in the message box once you press "Start", if you haven't disabled it. PPT Section Indicator should select this slide automatically, but if this doesn't happen (read [step 1](#step-1-formatting) for more details), select it manually.

#### I've received the following message: "PPT Section Indicator requires at least two sections if 'Include slide markers' is not selected"

If your presentation has only one section and you don’t create slide markers, the add-in will not work. This is actually related to the way the add-in works, since it relies on grouping the elements it creates, and PowerPoint does not allow to group a single element. However, after further consideration, there seems to be no point in creating a progress tracker if you only have one section and you're not interested in tracking individual slides. Therefore, either select "Include slide markers" or create more sections in your document.

#### I've received the following message: "Your presentation has changed while PPT Section Indicator was working. Please, restart the process"

This happens if you’ve inserted or deleted slides or section into your presentation after the tool started working. This check is necessary to avoid inconsistencies.

#### I've received the following message: "Unexpected error. Did you delete any element generated by PPT Section Indicator? Please, restart the process."

This error is most commonly caused if you've deleted some of the elements created by this add-in (either the section textboxes or slide markers). However, if this is not the case for you, then something unknown to me has happened. The error dialog should print the exception message to help find what caused the error.

## Acknowledgements

Some of the icons used in the toolbar were taken from [Icons8](https://icons8.com/).
