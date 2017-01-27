# PPT Section Indicator

PPT Section indicator is a PowerPoint VSTO add-in that adds section indicators on presentations, adapting them dynamically to highlight the current section or slide. The motivation behind this add-in is to automate the rather boring task of copying/pasting and reformatting objects in every slide of a presentation just to create a simple progress tracker.

For this tool to work, your presentation must be arranged into sections, as shown below.
![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/sections.png)

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

* First of all, **backup your presentation** before running the tool. It’s quite certain that I didn’t find all the bugs in this tool, and you may end up losing some information. Use this tool on your own risk;
* This add-in uses the clipboard to process objects among slides. Any content that you were holding on the clipboard will be lost once you run it;
* Since this tool creates objects in the presentation, PowerPoint requires me to name them (this naming process happens every time you add something to your presentation, but PowerPoint picks some name automatically). To avoid clashes, I’ve decided to name each element starting with “SectionIndicator”, as shown in the selection pane below. If, by any chance, this tool complains about naming collisions and you’ve never run it before on your presentation, you managed to include an object in your presentation and name it starting with this reserved string. **Do not run the cleanup function of the tool**, as it will erase this object thinking that the tool itself created it in a previous run. Instead, find it on the selection pane and rename it.

![](https://github.com/fabio-muramatsu/PPTSectionIndicator/blob/master/doc/images/selection_pane.png )

## Settings
This add-in has a few straightforward settings:

**Include slide markers**: This checkbox alternates between the two modes of operation, as explained before in this document;

**Slides to edit**: Determines which slides to include when processing the presentation. This option is useful if there are slides are not to be included in the marker count (e.g., title slides, appendices, etc.). For instance, if the first slide of the presentation were the title, you could include slides 2 through 9 only. This affects the tool in the following way:

* Slide 1 won’t be processed. In other words, the progress indicator won’t be added in the first slide;
* The “Intro” section will have only two slides, which will impact the number of markers created by the tool.

The syntax through which you specify slide ranges is very similar to the one used when printing a document. Contiguous ranges are denoted by a dash ("-"), and distinct slide ranges or single slides are separated by semi-colons (";"). Below are some examples of valid expressions:

* 2-9 ; 11
* 2-9 ; 11-20 ; 22
* 2 ; 3 ; 4 ; 10-15
