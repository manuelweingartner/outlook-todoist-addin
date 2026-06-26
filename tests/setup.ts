// Minimaler globaler Office-Mock. Einzelne Tests überschreiben Felder gezielt.
(global as any).Office = {
  CoercionType: { Html: "html", Text: "text" },
  AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  MailboxEnums: {
    AttachmentType: { File: "file", Item: "item", Cloud: "cloud" },
    AttachmentContentFormat: { Base64: "base64", Url: "url", Eml: "eml" },
  },
  context: {},
};
