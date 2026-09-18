# Saved attachment paths in remote deployments

`getMessage` and `getMessages` with `saveAttachments: true` save files in
**Thunderbird's** filesystem. The returned `attachments[].filePath` is not a
download to the MCP client. Saved attachment metadata includes
`filePathScope: "thunderbird-server"` and `filePathNote` to make this boundary
visible when reading tool results.

Do not run `cat`, `base64`, or another client-local file command against that
path from a separate agent container. A shared filesystem must be explicitly
configured before local access is possible. Otherwise obtain the original
bytes through an available transfer mechanism or ask for the original file.
Never fabricate Base64 or silently omit a requested attachment. These hints
do not introduce a file-download API or a shared volume.
