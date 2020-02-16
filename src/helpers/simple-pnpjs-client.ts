import { Group, User } from '@microsoft/microsoft-graph-types';
import { graph as graphClient } from '@pnp/graph-commonjs';
import { BearerTokenFetchClient } from '@pnp/nodejs-commonjs';

export class SimplePnPJsClient {
    
    private token: any;

    constructor(token: any) {
        if (!token || !token.trim()) {
            throw new Error('SimpleGraphClient: Invalid token received.');
        }

        this.token = token;

        graphClient.setup({
            graph: {
                fetchClientFactory: () => {
                    return new BearerTokenFetchClient(this.token);
                }
            }
        });
    }

    /**
     * Check if a user exists
     * @param {string} emailAddress Email address of the email's recipient.
     */
    public async userExists(emailAddress: string): Promise<User> {
        if (!emailAddress || !emailAddress.trim()) {
            throw new Error('SimplePnPjsClient.userExists(): Invalid `emailAddress` parameter received.');
        }
        return await graphClient.users.getById(emailAddress).get();
    }

    /**
     * Check if an alias is in use
     * @param {string} alias Alias for the group/site.
     */
    public async aliasExists(alias: string): Promise<Group> {
        if (!alias || !alias.trim()) {
            throw new Error('SimplePnPjsClient.aliasExists(): Invalid `alias` parameter received.');
        }
        return await graphClient.groups.filter(`mailNickname eq '${alias}' or displayName eq '${alias}'`).get();
    }
}
