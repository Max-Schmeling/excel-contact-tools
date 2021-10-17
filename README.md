# Excel Contact Tools

This Excel add-in was initially used in a German DAX-30 corporate to save its employees time when working with contact details (e.g. that of your co-workers) in the Active Directory. The tool has been saving hundreds of - if not more - working hours since 2019. I figured it can be used externally as well to help more people speed up certain monotonous tasks... so here we go :)


# How can the add-in help me?

The add-in assumes that your spread sheet contains contact details of the people who you are working with (e-mail address, full names, last names, first names, etc.). It also assumes that these contact details are managed in an Active Directory. If that is the case the add-in can do the following things for you:

1. **Show contact:** Show all other relevant contact details of the e-mail address or name in the active cell
2. **Convert contacts:** Convert multiple cells containing e-mail addresses or names into other contact details like phone numbers etc.
3. **Search address book**: Like Outlook's address book, but better ;)

These features will subsequently be explained. When the add-in is installed it will add a tab to Excel as shown in the following screenshot:

<p align="center">
  <img width="100%" src="/screenshots/ribbon-tab.png">
</p>


### Show contact (business card)
Show all contact details of the e-mail address or name in the active cell of your spread sheet. For example, if the cell you selected in your spreadsheet contains an e-mail address, a window will pop up displaying all contact details for that user. Clicking the `+`-button next to an entry will paste that entry into the active cell.

<p align="center">
  <img width="100%" src="/screenshots/business_card.png">
</p>

### Convert contacts
Select a range of contact details in your spread sheet and select a contact detail which you would like to have instead of the existing contact detail. For example, if your range contains e-mail addresses and names the tool will automatically convert all of those into phone numbers or street addresses or last names or whatever you selected. 

<p align="center">
  <img width="100%" src="/screenshots/convert_contacts.png">
</p>

If there are cases where the contact information in the selected cell(s) does not uniquely identify exactly one user - this can be the case when a cell contains a name and there are multiple people with the same name - the add-in will prompt you with a list of all users it has found allowing you to select the one you wish.

### Search Address book
Open a search bar and search for users in the Active Directory and list their contact details. Double-clicking an entry will open its business card.

<p align="center">
  <img width="100%" src="/screenshots/address_book.png">
</p>


# What are the technical requirements?

The entire thing will only work if your company's IT-infrastructure uses Active Directory to manage users... and you need access to that Active Directory. If you do not use Active Directory or are not connected to one, this add-in will be useless to you. The unavailability of the Active Directory will reveal itself through annoyingly long (30 seconds at max) Excel-freezes when attempting to run one of the tools.

# Compatibility

The add-in only works for local installations of Microsoft Excel, i.e., this add-in does not work with a browser-version of Microsoft Office.

# Installation

1. Download the file named *excel-contact-tools-addin.xlam*
2. Install the addin-file according to the instructions for **xlam**-addins. Head to the [official Microsoft website](https://support.microsoft.com/en-us/office/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) or simply ask google *How to install xlam-addins in Excel?*
4. Restart Excel
5. A new ribbon named *Contacts* should have appeared
6. If a yellow banner appears, you will have to click `Enable content` first before the add-in becomes usable to enable macros.

<p align="center">
  <img width="100%" src="/screenshots/enable_macros.png">
</p>

# Support

Since the project is no longer updated and maintained, any questions have to be answered using the available source code.
