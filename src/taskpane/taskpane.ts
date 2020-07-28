import { BrowserInfo } from "./browserInfo";

/* global document, Office */

Office.onReady(info => {
  const browserInfo = new BrowserInfo();
  const elementAppBody = document.getElementById("app-body");
  const elementBrowserName = document.getElementById("browserName");
  const elementBrowserVersion = document.getElementById("browserVersion");
  const elementHostAppName = document.getElementById("hostAppName");
  const elementHostAppType = document.getElementById("hostAppType");
  const elementSideloadMessage = document.getElementById("sideload-msg");

  if (browserInfo) {
    elementBrowserName.innerText = browserInfo.Name;
    elementBrowserVersion.innerText = browserInfo.Version;

    const elementBrowserLogo = getBrowserLogoElement(browserInfo);

    if (elementBrowserLogo) {
      elementBrowserLogo.style.display = "block";
    }
  }

  if (info.host) {
    elementHostAppName.innerText = getHostAppName(info.host);
    elementHostAppType.innerText = getHostAppType(info.platform);

    elementSideloadMessage.style.display = "none";
    elementAppBody.style.display = "flex";
  }
});

function getBrowserLogoElement(browserInfo: BrowserInfo): HTMLElement | undefined {
  const elementName = getBrowserLogoElementName(browserInfo);

  if (elementName) {
    return document.getElementById(elementName);
  }

  return undefined;
}

function getBrowserLogoElementName(browserInfo: BrowserInfo): string | undefined {
  switch (browserInfo.Name) {
    case "Chrome":
    case "Electron":
    case "Firefox":
    case "Safari":
      return `${browserInfo.Name.toLowerCase()}Logo`;
    case "Microsoft Edge":
      return isOldEdge(browserInfo.Version) ? "edgeOldLogo" : "edgeLogo";
    case "Internet Explorer":
      return "ieLogo";
    default:
      return undefined;
  }
}

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

function getHostAppType(platformType: Office.PlatformType) {
  switch (platformType) {
    case Office.PlatformType.Android:
      return "Android";
    case Office.PlatformType.Mac:
      return "Mac";
    case Office.PlatformType.OfficeOnline:
      return "Office Online";
    case Office.PlatformType.PC:
      return "Windows";
    case Office.PlatformType.Universal:
      return "Windows (Universal)";
    case Office.PlatformType.iOS:
      return "iOS";
  }
}

function isOldEdge(version: string) {
  if (version) {
    const majorVersion: number = parseInt(version.split(".")[0], 10);

    return majorVersion < 50;
  }

  return false;
}
