import { SimpleGraphClient } from './simple-graph-client';
import { SimplePnPJsClient } from './simple-pnpjs-client';

export class GraphHelper {
    /**
     * Let's the user see if the user exists using GraphClient
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     * @param {string} emailAddress The email address of the user.
     */
    public static async userExists(tokenResponse: any, emailAddress: string): Promise<boolean> {
        if (!tokenResponse) {
            throw new Error('GraphHelper.userExists(): `tokenResponse` cannot be undefined.');
        }
        const client = new SimpleGraphClient(tokenResponse.token);
        const user = await client.userExists(emailAddress);
        return user !== undefined;
    }

    /**
     * Let's the user see if the user exists using PnPJs
     * @param {TokenResponse} tokenResponse A response that includes a user token.
     * @param {string} emailAddress The email address of the user.
     */
    public static async userExistsPnP(tokenResponse: any, emailAddress: string): Promise<boolean> {
        if (!tokenResponse) {
            throw new Error('GraphHelper.userExists(): `tokenResponse` cannot be undefined.');
        }
        const client = new SimplePnPJsClient(tokenResponse.token);
        const user = await client.userExists(emailAddress);
        return user !== undefined;
    }
}
