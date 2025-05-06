# ArcGIS_Storymap_Creator
Converts a DOCX document to a storymap.

The easiest way is to download the whole ArcGIS Pro project folder (Storymap_creator_publish). Tested in ArcGIS Pro 3.4.3. (You can see script and config file code directly in the root folder, but they are also included in the project folder /_data/)

After opening the project, you can find the script ready to use in the project toolbox.

You need to fill out the connection to the server - URL, username, and password. If you will be making multiple Storymaps, you can use the config file in the data folder to save the connection properties and just load it from there. 

There are also two DOCX files in /_data/ folder that were used as tests.

Supported things in DOCX

Basic formatting that is also in Storymap - Bold, Italic, Subscript, Superscript, crossed, font color, plus underlined, which is not in Storymap but can actually be added. Font color match directly from DOCX - not theme colors.

All storymaps paragraph types - normal, heading 1-3, quote (matches to word style "quote"), code (matches to word style code - also tries to match language of code automatically, but best to check)

Lists are added automatically - Storymap has two limitations. Nested lists need to be of the same type, so the script matches nested items to the type of main list. In Storymap, there can only be two levels - so Storymap makes 3+ levels into level two and adds "---" per level before text to simulate deeper levels.

Images are uploaded, and if floating inside text, they do so in Storymap too. If there is a test with word style "caption" under it - or if caption is assigned to the image by right clicking the image and adding caption, the Caption is added to Storymap. I think it is better to add images later in better quality - Word makes images small automatically. In Storymap, you can add big images, and Storymap makes them small on the fly based on how they are set. So, for example image added from Word can never be full screen in Storymap - because it is just too small.
