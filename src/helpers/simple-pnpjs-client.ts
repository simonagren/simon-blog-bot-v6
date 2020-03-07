import { Group } from '@microsoft/microsoft-graph-types';
import { graph as graphClient } from '@pnp/graph-commonjs';
import { BearerTokenFetchClient } from '@pnp/nodejs-commonjs';

export class SimplePnPJsClient {
    
    private token: any;

    constructor(token: any) {
        if (!token || !token.trim()) {
            throw new Error('SimplePnPJsClient: Invalid token received.');
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
     * Check if an alias is in use
     * @param {string} alias Alias for the group/site.
     */
    public async aliasExists(alias: string): Promise<boolean> {
        if (!alias || !alias.trim()) {
            throw new Error('SimplePnPjsClient.aliasExists(): Invalid `alias` parameter received.');
        }
        try {
            const group: Group[] = await graphClient.groups.filter(`mailNickname eq '${alias}' or displayName eq '${alias}'`)();
            return group.length > 0;
        } catch (error) {
            return false;
        }
    }
}
