import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { getSP } from "../../../common/PnP/PnPConfig";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

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

export class CollaborationAudienceTargetService {
    public static async isCurrentUserInAudience(
        context: WebPartContext,
        audienceIds: string[]
    ): Promise<boolean> {
        if (!audienceIds || audienceIds.length === 0) {
            return true;
        }

        try {
            console.log("Audience IDs to check:", audienceIds);
            const isAdmin = await this.isCurrentUserAdmin(context);
            if (isAdmin) {
                console.log("Current user is an admin - bypassing audience checks");
                return true;
            }
            const currentUser = {
                loginName: context.pageContext.user.loginName?.toLowerCase(),
                email: context.pageContext.user.email?.toLowerCase(),
                groups: await this.getUserGroups(context)
            };

            console.log("Current user info:", {
                login: currentUser.loginName,
                email: currentUser.email,
                groupIds: currentUser.groups.map(g => g.Id.toString())
            });

            // Check each audience ID for matches
            for (const audienceId of audienceIds) {
                const normalizedId = audienceId?.toLowerCase()?.trim();
                if (!normalizedId) continue;

                // 1. Check direct user match (login or email)
                if (normalizedId === currentUser.loginName ||
                    normalizedId === currentUser.email) {
                    console.log(`Direct match found with audience ID: ${audienceId}`);
                    return true;
                }

                // 2. Check group membership
                const isGroupMatch = currentUser.groups.some(
                    group => group.Id.toString() === normalizedId
                );
                if (isGroupMatch) {
                    console.log(`Group membership match found with audience ID: ${audienceId}`);
                    return true;
                }
            }

            console.log("No audience matches found");
            return false;
        } catch (error) {
            console.error("Error in audience targeting check:", error);
            return false; // Fallback: don't show if check fails
        }
    }

    public static async isCurrentUserAdmin(context: WebPartContext): Promise<boolean> {
        try {
            const admins = await this.fetchAdmins(context);
            const currentUserLogin = context.pageContext.user.loginName?.toLowerCase();
            const currentUserEmail = context.pageContext.user.email?.toLowerCase();

            // Flatten all members from all groups
            const allMembers = admins.flatMap((group:any) => group.members);

            return allMembers.some((admin:any) =>
                admin.loginName?.toLowerCase() === currentUserLogin ||
                admin.email?.toLowerCase() === currentUserEmail
            );
        } catch (error) {
            console.error("Error checking admin status:", error);
            return false;
        }
    }

    private static getGraphClient(context: WebPartContext) {
        return Client.init({
            authProvider: async (done) => {
                const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
                const token = await tokenProvider.getToken("https://graph.microsoft.com");
                done(null, token);
            }
        });
    }

    // Helper: Fetch AAD group objectId by email
    public static async fetchAADGroupObjectId(groupEmail: string, context: WebPartContext): Promise<string | null> {
        const client = this.getGraphClient(context);
        const result = await client.api(`/groups?$filter=mail eq '${groupEmail}'`).get();
        console.log(`Fetched AAD group objectId for email ${groupEmail}:`, result);

        return result.value && result.value.length > 0 ? result.value[0].id : null;
    }

    // Helper: Fetch AAD group members by objectId
    public static async fetchAADGroupMembers(groupObjectId: string, context: WebPartContext): Promise<any[]> {
        const client = this.getGraphClient(context);
        const members = await client.api(`/groups/${groupObjectId}/members`).get();
        console.log(`Fetched ${members.value.length} members from AAD group ${groupObjectId}`);

        return members.value;
    }

    // Helper: Fetch SharePoint group members by groupId
    public static async fetchSPGroupMembers(groupId: number, context: WebPartContext): Promise<any[]> {
        try {
            const sp = getSP(context);
            const users = await sp.web.siteGroups.getById(groupId).users();
            console.log(`Fetched ${users.length} members from SharePoint group ${groupId}`);
            return users;
        } catch (error) {
            // If not found, return empty
            return [];
        }
    }

    // Main: Fetch all admins (users, SP groups, AAD groups)
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
                if (item.Admin && item.Admin.Id) {
                    // Try SharePoint group first
                    const spMembers = await this.fetchSPGroupMembers(item.Admin.Id, context);
                    if (spMembers.length > 0) {
                        groups.push({
                            groupName: item.Admin.Title,
                            groupType: "SharePoint",
                            members: spMembers.map((user: any) => ({
                                id: user.Id,
                                loginName: user.LoginName,
                                email: user.Email || "",
                                displayName: user.Title,
                            }))
                        });
                    } else if (item.Admin.EMail) {
                        // Try as Azure AD group by email
                        const groupObjectId = await this.fetchAADGroupObjectId(item.Admin.EMail, context);
                        if (groupObjectId) {
                            const aadMembers = await this.fetchAADGroupMembers(groupObjectId, context);
                            groups.push({
                                groupName: item.Admin.Title,
                                groupType: "AAD",
                                members: aadMembers.map((user: any) => ({
                                    id: user.id,
                                    loginName: user.userPrincipalName || user.mail,
                                    email: user.mail || "",
                                    displayName: user.displayName,
                                }))
                            });
                        } else {
                            // If not a group, treat as direct user admin
                            groups.push({
                                groupName: item.Admin.Title,
                                groupType: "User",
                                members: [{
                                    id: item.Admin.Id,
                                    loginName: item.Admin.Name,
                                    email: item.Admin.EMail || "",
                                    displayName: item.Admin.Title,
                                }]
                            });
                        }
                    } else {
                        // If not a group and no email, treat as direct user admin
                        groups.push({
                            groupName: item.Admin.Title,
                            groupType: "User",
                            members: [{
                                id: item.Admin.Id,
                                loginName: item.Admin.Name,
                                email: item.Admin.EMail || "",
                                displayName: item.Admin.Title,
                            }]
                        });
                    }
                }
            }
            return groups;
        } catch (error) {
            console.error("Error fetching admin groups with members:", error);
            return [];
        }
    }

    public static async getUserGroups(context: WebPartContext): Promise<{ Id: string | number }[]> {
        try {
            const response = await context.spHttpClient.get(
                `${context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups`,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Failed to fetch user groups: ${response.statusText}`);
            }

            const data = await response.json();
            return data.value || [];
        } catch (error) {
            console.error("Error fetching user groups:", error);
            return [];
        }
    }
}