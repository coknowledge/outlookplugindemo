Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("insert-link").onclick = insertLink;
    }
});

function insertLink() {
    const url = "https://www.coknowledge.nl";
    
    Office.context.mailbox.item.body.setSelectedDataAsync(
        `<p><a href="${url}" target="_blank">DEMO URL</a></p>`,
        {
            coercionType: Office.CoercionType.Html
        },
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to insert link:", result.error);
            } else {
                // Show success message
                document.getElementById("insert-link").innerHTML = "✓ Link Added!";
                setTimeout(() => {
                    document.getElementById("insert-link").innerHTML = 
                        '<span class="ms-Button-label">Add DEMO URL</span>';
                }, 2000);
            }
        }
    );
}