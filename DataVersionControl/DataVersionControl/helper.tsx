import * as React from "react";
import { IInputs } from "./generated/ManifestTypes";

export function useNoteAttachment(objectId:string, context: ComponentFramework.Context<IInputs>) {
    const [noteAttachment, setNoteAttachment] = React.useState<any>();
    React.useEffect(() => {
        const fetchData = {
            "objectid": objectId.replace(/[{}]/g, ''),
            "isdocument": "1"
          };
        let fetchXml = [
        "<fetch>",
        "  <entity name='annotation'>",
        "    <filter>",
        "      <condition attribute='objectid' operator='eq' value='", fetchData.objectid/*00000000-0000-0000-0000-000000000000*/, "'/>",
        "      <condition attribute='isdocument' operator='eq' value='", fetchData.isdocument/*1*/, "'/>",
        "      <condition attribute='documentbody' operator='not-null'/>",
        "    </filter>",
        "  </entity>",
        "</fetch>"
        ].join("");
        fetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
        context.webAPI.retrieveMultipleRecords("annotation", fetchXml).then((response: any)=>{
            setNoteAttachment(response.entities);
        });
    }, []);
    return noteAttachment;
}
