import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, Web, SiteUsers, SiteGroups } from "@pnp/sp/presets/all";

export class SPService {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async getCurrentSiteUsers(siteUrl: string): Promise<any> {
        try {
            let currentWeb = Web(siteUrl)
            let users: any[] = await currentWeb.siteUsers();
            return users;
        } catch (err) {
            Promise.reject(err);
        }
    }

    public async getCurrentSiteGroups(siteUrl: string): Promise<any> {
        try {
            let currentWeb = Web(siteUrl)
            const groups: any[] = await currentWeb.siteGroups()
            return groups;
        }
        catch (err) {
            Promise.reject(err);
        }
    }
}
