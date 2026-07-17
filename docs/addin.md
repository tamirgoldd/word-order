# Sideload the Word add-in

The add-in is an alpha developer build. First build the repository, then serve
`packages/addin/dist` over HTTPS and replace `https://localhost:3001` in
`packages/addin/manifest.xml` with that origin.

In Word, open the add-ins dialog, choose **My Add-ins**, then **Upload My
Add-in**, and select the manifest. Microsoft 365 administrators can deploy the
same manifest through Integrated Apps after choosing a stable hosted origin.

The task pane scans before applying and warns users to save a backup. Word's
OOXML insertion APIs differ by host version, so the add-in reports unsupported
hosts rather than attempting a partial repair.
