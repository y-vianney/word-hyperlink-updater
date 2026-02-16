# Word Hyperlink Updater Add-in

## How to Run

1.  **Serve the files**:
    You need to serve these files at `https://localhost:3000`.
    You can use `http-server` with SSL:
    ```bash
    npx http-server -S -p 3000
    ```
    *Note: You may need to generate/trust a self-signed certificate.*

2.  **Sideload into Word**:
    - Go to **Word on the Web** (or Desktop).
    - **Insert** > **Add-ins**.
    - **Manage My Add-ins** > **Upload My Add-in**.
    - Select the `manifest.xml` file.

3.  **Use**:
    - The "Hyperlink Updater" button (or similar) should appear in the Home tab, or a task pane will open.
    - Enter your SharePoint URL (e.g., `https://contoso.sharepoint.com/sites/mysite/`).
    - Click **Update Hyperlinks**.
