# Excel to Powerpoint VBA Example

## Purpose
This project was created as an example of an Excel to Powerpoint workflow using VBA. Typically one might create PowerPoint slides from data in Excel and by using VBA we can **automate** the process, saving time and ensuring consistency.

By investing the time up-front in creating a PowerPoint template and writing some VBA code you can save yourself from hours/day of tedious work.

## Using the code
1. Download the "data_file.xlsm" and "presentation_template.pptx"
2. Open the "data_file.xlsm" file and click the "Create Presentation" button on the ribbon.
    * You will be prompted to select a PowerPoint file to use as a template for the presentation you're building. Select the "presnentation_template.pptx" file.
    * Note: Since the "Create Presentation" button is in the **ribbon** there's also a **keyboard shortcut**, namely ALT, H, Y. This (and the fact that ribbon buttons cannot be accidentally deleted by the end user) is the main reason I prefer to add buttons to the ribbon instead of using form controls or ActiveX controls.
3. Profit!
    * You just created a PowerPoint presentation with custom chart formatting in a few seconds with a few clicks.

## Extensions
This code is by no means meant to be a "complete" example of the capabilities of VBA or any "Excel-to-Powerpoint" workflows, but it does serve as a basis that can easily be extended to your requirements. Such extensions may include (but are not limited to):

* Additional chart types
* Integration with data from MS Access or another database system
* Adding the ability to use custom themes/colour palettes for your organisation or for a particular client
* Adding a "configuration" worksheet or user form to dynamically select which data to pull and which slides to which to push that data