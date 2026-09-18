# Saved attachment paths in remote deployments

`getMessage` and `getMessages` with `saveAttachments: true` save files in
**Thunderbird's** filesystem. The returned `attachments[].filePath` is not a
download to the MCP client. Saved attachment metadata includes
`filePathScope: "thunderbird-server"` and `filePathNote` to make this boundary
visible when reading tool results.

Do not run `cat`, `base64`, or another client-local file command against that
path from a separate agent container. A shared filesystem must be explicitly
configured before local access is possible.

To reuse an existing email attachment across containers:

1. Call `getMessage` with the same `messageId` and `folderPath`, plus
   `rawSource: true`. The response's `rawSource` contains the full message,
   including MIME attachment parts; no shared filesystem is required.
2. Parse that source with a MIME parser and select the intended attachment.
   Decode its `Content-Transfer-Encoding` to obtain the original bytes.
3. Base64-encode those exact bytes and pass
   `{name, contentType, base64}` to the sending tool. Keep the original
   filename and MIME type. Do not attach the entire raw email as the PDF.

Perform extraction and encoding programmatically; never manually reconstruct,
shorten or invent Base64. Check that the response is complete before parsing.
The raw source string represents original bytes as Latin-1; use Latin-1
when converting it back to bytes for a byte-oriented MIME parser.

`getMessages(rawSource: true)` provides the same raw-source route for a batch.
Raw source access requires a local/offline message copy; an uncached IMAP
message may fail. Report that specific failure or obtain an accessible
original rather than silently omitting the requested attachment.
These hints use the existing raw-source API and add no new transfer tool.

