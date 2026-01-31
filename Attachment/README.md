# Custom Power Apps Attachment Component

A highly customisable, modern replacement for the standard Power Apps attachment control. This component allows developers to maintain a sleek UI while leveraging native attachment functionality and Power Automate for file processing.

## Features

* **Fully Themeable**: Custom properties for colours, fills, and border radius.
* **Dynamic SVG Support**: Change the icon colour dynamically via the `Color` property using SVG string manipulation.
* **Adaptive Layout**: Supports both Horizontal and Vertical label directions.
* **Smart File Handling**: Auto-generates GUIDs for session tracking and handles duplicate filename conflicts (e.g. `file (1).png`).
* **Flow-Ready**: Pre-processes attachments into base64 strings to be passed into Power Automate.

## Installation

* Copy the `YAML` code in `Attachment.yml` into the Components tab.


## Component Properties

|Property|Type|Description|
|:-|:-|:-|
|


### Component Variables


## Power Automate Flow Integration

When triggering your flow, pass the `colAttachments` as a JSON string. In my example flow, I have set it up so that it can be used dynamically in any SharePoint Site and Library.

The flow trigger should have the following parameters:

* Site: SharePoint Site URL.
* Library: Display Name of the Library (the flow will use the display name).
* Content: The collection containing the files to be attached from the component.
* Metadata: Additional metadata to be added to the attached files. Ensure that the columns exist in the target library.

The flow also needs to have the **Respond to a Power App or flow** action which returns a `response` as an output. This `reponse` is used in the component to refresh the data source and perform any additional logic when the flow is successful.

In the flow:

1. Use the `json()` expression to parse the `Content` and `Metadata` inputs.
1. Use an `Apply to each` loop to create each of the attachments to the target SharePoint Library.
1. Pass the following

## Technical Deep Dive

### The "Hack"

The component uses the native `Attachments` control layered with a 0% opacity fill and positioned so that it the native OS file picker is triggered. The look and feel is set up in a container behind the `Attachments`.

### Extracting Data from the Attachments Control

The component converts the blob URI of a file that is attached to a base64 string within the `OnAddFile` property:

```powerfx
ForAll(
    galFileContent_cmpAttachment.AllItems As attachmentFile,
    With(
        {
            tempID: Text(GUID()),
            contentFile: Index(Split(Substitute(JSON(attachmentFile.imgFileContent_cmpAttachment.Image, JSONFormat.IncludeBinaryData), """", ""), ","),2).Value,
            numDuplicates: CountRows(Filter(colAttachments, Name = attachmentFile.Name)),
            //.test: First(Split(attachmentFile.Name, "." & Last(Split(attachmentFile.Name, ".").Value))),
            fileName: First(Split(attachmentFile.Name, ".")).Value,
            fileExtension: Last(Split(attachmentFile.Name, ".")).Value
        },
        Collect(
            colAttachments,
            {
                Name: If(!IsBlank(LookUp(colAttachments, Name = attachmentFile.Name)), $"{fileName} ({numDuplicates}).{fileExtension}", attachmentFile.Name),
                Status: "Unsaved",
                ID: tempID,
                FileContent: ""
            }
        );
        //Notify(LookUp(colAttachments, ID = tempID).ID);
        // TODO: IF OVER THE ATTACHMENT SIZE --> DONT ADD
        Patch(
            colAttachments,
            LookUp(colAttachments, ID = tempID),
            {
                FileContent: contentFile
            }
        );
    );
);
```

## Contributing



