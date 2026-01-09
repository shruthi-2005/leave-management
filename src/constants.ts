import { WebPartContext } from "@microsoft/sp-webpart-base";

/** Centralized SharePoint List Names */
export const LIST_NAMES = {
  Tasks: "Tasks",
  LeaveRequests:"LeaveRequests",
  LeaveInformations:"LeaveInformations",
  Holidays:"Holidays"
};

/** Get Site URL (Works in Workbench, Teams, and Site page) */
export const getSiteUrl = (context: WebPartContext): string => {
  try {
    // 1️⃣ Try standard web URL
    const siteUrl = context?.pageContext?.web?.absoluteUrl;
    if (siteUrl && siteUrl.toLowerCase().includes("trainingportal")) {
      return siteUrl;
    }

    // 2️⃣ Try using site absolute URL (Teams sometimes returns only root web)
    const siteAbsoluteUrl = (context as any)?.pageContext?.site?.absoluteUrl;
    if (siteAbsoluteUrl && siteAbsoluteUrl.toLowerCase().includes("trainingportal")) {
      return siteAbsoluteUrl;
    }

    // 3️⃣ If Teams trimmed it to root (common bug), hardcode fallback
    return "https://elevix.sharepoint.com/sites/Trainingportal";
  } catch (error) {
    console.error("Error getting site URL:", error);
    return "https://elevix.sharepoint.com/sites/Trainingportal";
  }
}; 