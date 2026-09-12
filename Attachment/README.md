# Custom Power Apps Attachment Component

A highly customisable, modern replacement for the standard Power Apps attachment control. This component allows developers to maintain a sleek UI while leveraging native attachment functionality and Power Automate for file processing.

## Features

* **Fully Themeable**: Custom properties for colours, fills, and border radius.
* **Dynamic SVG Support**: Change the icon colour dynamically via the `Color` property using SVG string manipulation.
* **Adaptive Layout**: Supports both Horizontal and Vertical label directions.
* **Smart File Handling**: Auto-generates GUIDs for session tracking and handles duplicate filename conflicts (e.g. `file (1).png`).
* **Flow-Ready**: Pre-processes attachments into base64 strings to be passed into Power Automate.

## Getting Started

* Copy the `YAML` code in `Attachment.yml` into the Components tab of your Power App or Component Library.
* Create a flow to upload the output of the component.

## Solution Example

A solution containing an example of the component is provided in `Attachments_1_0_0_0.zip`. The solution contains:

* An example app: **Attachments** which contains an upload on a button click and a upload on file attachment.
* An example flow: **UploadFiles** which is used to upload the files to a SharePoint Document Library and can include metadata columns. The flow can be used across different libraries.

## Component Properties

### Custom Properties Reference: `cmpAttachment`

| Property Internal Name | Display Name | Type | Data Type | Description |
| :--- | :--- | :--- | :--- | :--- |
| `AttachFile` | **Attach File** | Event | N/A | The primary event where you place your Power Automate Flow or Patch logic. |
| `Attachments` | **Attachments** | Output | Table | The resulting table of files (Name, ID, FileContent, etc.) after processing. |
| `BackgroundFill` | **Background Fill** | Input | Color | The primary fill color of the component container. |
| `BorderRadius` | **Border Radius** | Input | Number | Controls the corner rounding of the component. |
| `BorderStyle` | **Border Style** | Input | Text | The style of the border (Dashed, Solid, None). |
| `BorderThickness` | **Border Thickness** | Input | Number | The thickness of the outer container border. |
| `ClearAttachments` | **Clear Attachments** | Action | N/A | Resets the internal file collection (`colAttachments`). |
| `Color` | **Color** | Input | Color | The color of the text label and the SVG icon. |
| `DisabledBackgroundFill` | **Disabled Background Fill** | Input | Color | Background color used when the control is in View/Disabled mode. |
| `DisabledColor` | **Disabled Color** | Input | Color | Text and Icon color used when the control is in View/Disabled mode. |
| `DisplayMode` | **Display Mode** | Input | Text | Controls whether the component is in Edit, View, or Disabled mode. |
| `Font` | **Font** | Input | Text | The font family used for the label text. |
| `HoverFill` | **Hover Fill** | Input | Color | The color that appears when a user hovers over the control. |
| `Icon` | **Icon** | Input | Text | The SVG string used for the attachment icon. |
| `IconSize` | **Icon Size** | Input | Number | The height and width of the icon in pixels. |
| `LabelDirection` | **Label Direction** | Input | Text | Sets the layout to `LayoutDirection.Horizontal` or `Vertical`. |
| `Loading` | **Loading** | Output | Boolean | Becomes `true` while the component is processing binary data. |
| `MaxAttachments` | **Max Attachments** | Input | Number | The maximum number of files allowed (Internal to native control). |
|`MaxAttachmentSize` | **Maximum attachment size (in MB)** | Input | Number | The maximum size allowed for an upload. |
| `OnAddFile` | **On Add File** | Action | N/A | Internal action that executes the `AttachFile` event. |
| `PressedFill` | **Pressed Fill** | Input | Color | The color that appears when the control is clicked. |
| `RemoveFile` | **Remove File** | Action | N/A | Action to remove a file from the collection based on its GUID (`IDToDelete`). |
| `ResetOnAdd` | **Reset On Add** | Input | Boolean | If true, wipes the selection after every individual file upload. |
| `Size` | **Font Size** | Input | Number | The size of the "Upload File" text. |
| `Text` | **Text** | Input | Text | The display label (e.g., "Upload File", "Attach Invoice"). |
| `ValidateMaximumAttachments` | **Validate Maximum Attachments** | Output Function | Boolean | A function to check if adding a file exceeds a specific limit. |


### Component Variables

| Name | Type | Purpose |
| :--- | :--- | :--- |
| `gblLoadingData` | Boolean | A Boolean flag used to toggle the Loading output property. It is set to true while the component iterates through files and extracts Base64 strings, then false once finished. |
| `gblContent` | Boolean | Stores the raw Self.Attachments table from the hidden native attachment control. It serves as the source for the internal gallery (galFileContent_cmpAttachment) used to extract binary data. |
| `gblColor` | String | While not a Power App variable, this is used as a string placeholder within the SVG Icon property, allowing for dynamic color injection via the Substitute function. |
| `colAttachments` | Collection | The primary internal collection that stores the file metadata and Base64 content (Name, FileNameWithoutExtension, FileExtension, Status, ID, FileContent). |
| `colNumberExceededAttachments` | Collection| A temporary "error" collection used during the `OnAddFile` event in the attachment control to track files that were rejected because the `MaxAttachments` limit was reached.

## Power Automate Flow Integration

When triggering your flow, pass the `colAttachments` as a JSON string. In my example flow, I have set it up so that it can be used dynamically in any SharePoint Site and Library.

The flow trigger should have the following parameters:

* Site: SharePoint Site URL.
* Library: Display Name of the Library (the flow will use the display name).
* Content: The collection containing the files to be attached from the component.
* Metadata: Additional metadata to be added to the attached files. Ensure that the columns exist in the target library.

The flow also needs to have the **Respond to a Power App or flow** action which returns a `response` as an output. This `response` is used in the component to refresh the data source and perform any additional logic when the flow is successful.

In the flow:

1. Use the `json()` expression to parse the `Content` and `Metadata` inputs.
1. Use an `Apply to each` loop to create each of the attachments to the target SharePoint Library.
1. Pass the following

## Technical Deep Dive

### The "Hack"

The component uses the native `Attachments` control layered with a 0% opacity fill and positioned so that it the native OS file picker is triggered. The look and feel is set up in a container behind the `Attach

### Extracting Data from the Attachments Control

The component converts the blob URI of a file that is attached to a base64 string within the `OnAddFile` property of the attachment control:

```powerfx
// Resets the component state whenever this runs. 
Set(gblContent, Self.Attachments);
Set(gblLoadingData, true);
Clear(colNumberExceededAttachments);

With(
    {
        stagedAttachments: ForAll(
            galFileContent_cmpAttachment.AllItems As attachmentFile,
            With(
                {
                    tempID: Text(GUID()),
                    contentFile: Index(Split(Substitute(JSON(attachmentFile.imgFileContent_cmpAttachment.Image, JSONFormat.IncludeBinaryData), """", ""), ","),2).Value,
                    numDuplicates: CountRows(Filter(colAttachments, Name = attachmentFile.Name)),
                    fileNameWithoutExtension: Substitute(attachmentFile.Name, $".{Last(Split(attachmentFile.Name, ".")).Value}", ""),
                    fileExtension: Last(Split(attachmentFile.Name, ".")).Value
                },
                {

                    Name: If(
                        !IsBlank(LookUp(colAttachments, Name = attachmentFile.Name)), 
                        $"{fileNameWithoutExtension} ({numDuplicates}).{fileExtension}", 
                        attachmentFile.Name
                    ),
                    FileNameWithoutExtension: fileNameWithoutExtension,
                    FileExtension: fileExtension,
                    Status: "Unsaved",
                    ID: tempID,
                    FileContent: contentFile
                }
            )
        )
    },
    With(
        {
            allowedSlots: atcAttachment_cmpAttachment.MaxAttachments - CountRows(colAttachments),
            totalNew: CountRows(stagedAttachments)
        },
        Collect(
            colAttachments,
            FirstN(stagedAttachments, allowedSlots)
        );
        
        // If the user attached more than allowed, generate dummy error records for the notification
        With(
            {
                numberExceededAttachments: If(
                    totalNew > allowedSlots,
                    Collect(
                        colNumberExceededAttachments,
                        ForAll(Sequence(totalNew - allowedSlots), {})
                    )
                )
            },
            If(
                !IsEmpty(colNumberExceededAttachments),
                Notify($"{CountRows(colNumberExceededAttachments)} attachments not added. Maximum attachments reached.", NotificationType.Error, 4000)
            );
        );

    );
);

If(
    cmpAttachment.ResetOnAdd,
    cmpAttachment.OnAddFile();
    cmpAttachment.ClearAttachments();
);

Reset(Self);
Set(gblLoadingData, false);
```
The collection is loaded into a gallery within the component, which contains an image control that converts the file into a `base64` string. 

This is then passed into a Power Automate flow which performs the upload of the attached content into a document library.

**Example:**

```powerfx
Set(
    gblUploadFilesResponse, 
    UploadFiles.Run(JSON(cmpAttachment_Attachments.Attachments), JSON(nfAttachmentMetadata), "https://55nwjd.sharepoint.com/sites/ComponentsPlayground", "AttachmentLibrary");
);

If(
    Value(gblUploadFilesResponse.response) = 200,
    // Refresh your Data Source here
    Refresh(AttachmentLibrary);
    cmpAttachment_Attachments.ClearAttachments();
    Notify("Success", NotificationType.Success, 4000)
    ,
    Notify("Error", NotificationType.Error, 4000)
)
```

