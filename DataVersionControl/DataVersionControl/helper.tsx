import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";
import { IComboBoxOption, IDropdownOption } from "@fluentui/react";

export async function getNoteAttachment(neoProcessId:string, context: ComponentFramework.Context<IInputs>, option?: IDropdownOption) {
    //const [noteAttachment, setNoteAttachment] = React.useState<any>();
    let attachments: any = [];
        const fetchData = {
            "isdocument": "1",
            "filename": ".csv",
            "vel_type": option?.key.toString(),
            "vel_operationalchangerequestid": neoProcessId.replace(/[{}]/g, '')
          };
        let fetchXml = [
        "<fetch>",
        "  <entity name='annotation'>",
        "    <filter>",
        "      <condition attribute='documentbody' operator='not-null'/>",
        "      <condition attribute='isdocument' operator='eq' value='", fetchData.isdocument/*1*/, "'/>",
        "      <condition attribute='filename' operator='ends-with' value='", fetchData.filename/*.csv*/, "'/>",
        "    </filter>",
        "    <link-entity name='vel_import' from='vel_importid' to='objectid'>",
        "      <filter>",
        "        <condition attribute='vel_type' operator='eq' value='", fetchData.vel_type/*0*/, "'/>",
        "        <condition attribute='vel_operationalchangerequestid' operator='eq' value='", fetchData.vel_operationalchangerequestid/*1d7e6011-2e0c-4c06-9d3f-7997596a2419*/, "'/>",
        "      </filter>",
        "    </link-entity>",
        "  </entity>",
        "</fetch>"
        ].join("");
        fetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
        await context.webAPI.retrieveMultipleRecords("annotation", fetchXml).then(
            function success (response: any){
                attachments = response.entities;
                return true;
            },
            function error(e){
                throw e;
            }
        );
    return attachments;
}

export function generateComboBoxOptions (noteAttachments: any){
   const options:IComboBoxOption[] = noteAttachments.map((item: any)=>{
        const fileName = item.filename;
        const createdOn = item['createdon@OData.Community.Display.V1.FormattedValue'];
        return({key:item.annotationid, text:fileName + ' - ' + createdOn,})
    })
    return options;
}
