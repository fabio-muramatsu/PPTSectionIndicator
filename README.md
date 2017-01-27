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
