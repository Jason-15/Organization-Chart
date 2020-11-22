import { WebPartContext } from "@microsoft/sp-webpart-base";
export default class SPServices {
    private context;
    constructor(context: WebPartContext);
    getUserProperties(user: string): Promise<any>;
    /**
     * async GetUserProfileProperty
     * user:string
     */
    getUserProfileProperty(user: string, property: string): Promise<any>;
}
//# sourceMappingURL=SPServices.d.ts.map