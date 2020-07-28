import { BrowserInfo } from "./browserInfo";

/* global document, Office */

Office.onReady(info => {
    const browserInfo = new BrowserInfo();

    if (browserInfo)
    {
        document.getElementById("browserName").innerText = browserInfo.Name;
        document.getElementById("browserVersion").innerText = browserInfo.Version;
    }

    if (info.host) {
        document.getElementById("hostAppName").innerText = getHostAppName(info.host);

        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";    
    }
});


function getHostAppName(host: Office.HostType) {
    switch (host) {
        case Office.HostType.Access:
            return "Access";
        case Office.HostType.Excel:
            return "Excel";
        case Office.HostType.OneNote: 
            return "OneNote";
        case Office.HostType.Outlook:
            return "Outlook";
        case Office.HostType.PowerPoint:
            return "PowerPoint";
        case Office.HostType.Project: 
            return "Project";
        case Office.HostType.Word:
            return "Word";
        default:
            return "";
    }
}