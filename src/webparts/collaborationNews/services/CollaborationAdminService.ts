import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

// Types
interface SharePointUser {
  Id: number;
  LoginName: string;
  Email: string;
  Title: string;
}

interface AADUser {
  id: string;
  userPrincipalName?: string;
  mail?: string;
  displayName: string;
}

interface AdminGroupWithMembers {
  groupName: string;
  groupType: "SharePoint" | "AAD" | "User";
  members: {
    id: string;
    loginName: string;
    email: string;
    displayName: string;
  }[];
}

export class CollaborationAdminService {
  // Check if current user is admin (user, SP group, or AAD group)
  public static async isCurrentUserAdmin(context: WebPartContext): Promise<boolean> {
    try {
      const admins = await this.fetchAdmins(context);
      const currentUserLogin = context.pageContext.user.loginName?.toLowerCase();
      const currentUserEmail = context.pageContext.user.email?.toLowerCase();

      // Flatten all members from all groups
    // Replace this:
// const allMembers = admins.flatMap(group => group.members);

// With this:
const allMembers = admins.reduce(
  (acc: { loginName: string; email: string }[], group: AdminGroupWithMembers) => acc.concat(group.members),
  []
);

return allMembers.some(
  (admin: { loginName: string; email: string }) =>
    admin.loginName?.toLowerCase() === currentUserLogin ||
    admin.email?.toLowerCase() === currentUserEmail
);
    } catch (error) {
      console.error("Error checking admin status:", error);
      return false;
    }
  }

  // Fetch all admins (users, SP groups, AAD groups)
  public static async fetchAdmins(context: WebPartContext): Promise<AdminGroupWithMembers[]> {
    const groups: AdminGroupWithMembers[] = [];
    try {
      const response = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('AdminList')/items?$select=Admin/Id,Admin/Title,Admin/EMail,Admin/Name&$expand=Admin`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) throw new Error(`Failed to fetch admins: ${response.statusText}`);
      const data = await response.json();
      const items = data.value || [];

      for (const item of items) {
        const admin = item.Admin;
        if (admin && admin.Id) {
          // Try SharePoint group first
          const spMembers = await this.fetchSPGroupMembers(admin.Id, context);
          if (spMembers.length > 0) {
            groups.push({
              groupName: admin.Title,
              groupType: "SharePoint",
              members: spMembers.map((user: SharePointUser) => ({
                id: user.Id.toString(),
                loginName: user.LoginName,
                email: user.Email || "",
                displayName: user.Title,
              }))
            });
            console.log("Fetched SharePoint group:", admin.Title, "with members:", spMembers);
          } else if (admin.EMail) {
            // Try as Azure AD group by email
            const groupObjectId = await this.fetchAADGroupObjectId(admin.EMail, context);
            if (groupObjectId) {
              const aadMembers = await this.fetchAADGroupMembers(groupObjectId, context);
              groups.push({
                groupName: admin.Title,
                groupType: "AAD",
                members: aadMembers.map((user: AADUser) => ({
                  id: user.id,
                  loginName: user.userPrincipalName || user.mail || "",
                  email: user.mail || "",
                  displayName: user.displayName,
                }))
              });
            } else {
              // If not a group, treat as direct user admin
              groups.push({
                groupName: admin.Title,
                groupType: "User",
                members: [{
                  id: admin.Id.toString(),
                  loginName: admin.Name,
                  email: admin.EMail || "",
                  displayName: admin.Title,
                }]
              });
            }
          } else {
            // If not a group and no email, treat as direct user admin
            groups.push({
              groupName: admin.Title,
              groupType: "User",
              members: [{
                id: admin.Id.toString(),
                loginName: admin.Name,
                email: admin.EMail || "",
                displayName: admin.Title,
              }]
            });
             console.log("Fetched admin groups with members:", groups);
          }
        }
      }
      return groups;
   
    } catch (error) {
      console.error("Error fetching admin groups with members:", error);
      return [];
    }
  }

  // Fetch SharePoint group members by groupId
  public static async fetchSPGroupMembers(groupId: number, context: WebPartContext): Promise<SharePointUser[]> {
    try {
      const response = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/sitegroups(${groupId})/users`,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) return [];
      const data = await response.json();
      return data.value || [];
    } catch {
      return [];
    }
  }

  // Fetch AAD group objectId by email
  public static async fetchAADGroupObjectId(groupEmail: string, context: WebPartContext): Promise<string | null> {
    const client = this.getGraphClient(context);
    const result = await client.api(`/groups?$filter=mail eq '${groupEmail}'`).get();
    return result.value && result.value.length > 0 ? result.value[0].id : null;
  }

  // Fetch AAD group members by objectId
  public static async fetchAADGroupMembers(groupObjectId: string, context: WebPartContext): Promise<AADUser[]> {
    const client = this.getGraphClient(context);
    const members = await client.api(`/groups/${groupObjectId}/members`).get();
    return members.value || [];
  }

  // Get Microsoft Graph client
  private static getGraphClient(context: WebPartContext) {
    return Client.init({
      authProvider: async (done) => {
        const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
        const token = await tokenProvider.getToken("https://graph.microsoft.com");
        done(null, token);
      }
    });
  }
}