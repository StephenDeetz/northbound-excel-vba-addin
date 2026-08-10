# Installing Northbound.xlam

Currently tested on English (US) Excel for Windows. Other locales may work but are unverified.

Northbound is an Excel VBA add-in. Follow these steps once; after that it loads automatically every time Excel starts.

## 1. Save the file

1. Create the folder: `Documents\Northbound Add-in`
2. Save `Northbound.xlam` into that folder.

## 2. Unblock the file

Windows blocks files downloaded from the internet or received by email. If you skip this step, the add-in may load but its macros will be disabled.

1. Open File Explorer and navigate to `Documents\Northbound Add-in`.
2. Right-click `Northbound.xlam` and choose **Properties**.
3. On the **General** tab, at the bottom, look for a **Security** section with the message *"This file came from another computer..."*.
4. Check the **Unblock** box.
5. Click **OK**.

If there is no Security section, the file is already unblocked. Continue.

## 3. Add the folder as a Trusted Location

Since a 2016 Office security update, add-ins outside a Trusted Location are blocked.

1. Open Excel with a blank workbook.
2. **File** > **Options** > **Trust Center** > **Trust Center Settings...**
3. Select **Trusted Locations** in the left pane.
4. Click **Add new location...**
5. Click **Browse...** and select `Documents\Northbound Add-in`.
6. Click **OK** to close each dialog.

## 4. Register the add-in with Excel

1. In Excel: **File** > **Options** > **Add-ins**.
2. At the bottom, set **Manage:** to **Excel Add-ins**.
3. Click **Go...**
4. In the Add-Ins dialog, click **Browse...**
5. Navigate to `Documents\Northbound Add-in`, select `Northbound.xlam`, click **OK**.
6. Confirm **Northbound** appears in the list with its checkbox ticked.
7. Click **OK**.

The add-in is now installed. It will load automatically when Excel starts.

## Updating

1. Close Excel.
2. Replace `Northbound.xlam` in `Documents\Northbound Add-in` with the new version.
3. Reopen Excel. If the new file was downloaded, repeat step 2 (Unblock).

## Uninstalling

To disable (keep the file, stop loading):

1. **File** > **Options** > **Add-ins** > **Manage: Excel Add-ins** > **Go...**
2. Untick **Northbound** and click **OK**.

To fully remove:

1. Disable as above.
2. Close Excel.
3. Delete `Documents\Northbound Add-in\Northbound.xlam`.
4. Reopen Excel. In the Add-Ins dialog, click **Northbound** - Excel will prompt to remove the missing reference. Click **Yes**.
