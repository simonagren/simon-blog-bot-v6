import { Client } from '@microsoft/microsoft-graph-client';
import { Group, User } from '@microsoft/microsoft-graph-types';

export class SimpleGraphClient {
    
    private token: string;
    private graphClient: Client;

    constructor(token: any) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this.token = token;

        // Get an Authenticated Microsoft Graph client using the token issued to the user.
        this.graphClient = Client.init({
            authProvider: (done) => {
                done(null, this.token); // First parameter takes an error if you can't get an access token.
            }
        });
    }

    /**
     * Check if a user exists
     * @param {string} emailAddress Email address of the email's recipient.
     */
    public async userExists(emailAddress: string): Promise<User> {
        if (!emailAddress || !emailAddress.trim()) {
            throw new Error('SimpleGraphClient.userExists(): Invalid `emailAddress` parameter received.');
        }
        return await this.graphClient
            .api(`/users/${emailAddress}`)
            .get().then((res: User) => {
                return res;
            });
    }

    /**
     * Check if an alias is in use
     * @param {string} alias Alias for the group/site.
     */
    public async aliasExists(alias: string): Promise<Group> {
        if (!alias || !alias.trim()) {
            throw new Error('SimpleGraphClient.aliasExists(): Invalid `alias` parameter received.');
        }
        return await this.graphClient
            .api('/groups/').filter(`mailNickname eq '${alias}' or displayName eq '${alias}'`)
            .get().then((res: Group) => {
                return res;
            });
    }
}
