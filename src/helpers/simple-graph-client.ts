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
    public async userExists(emailAddress: string): Promise<boolean> {
        if (!emailAddress || !emailAddress.trim()) {
            throw new Error('SimpleGraphClient.userExists(): Invalid `emailAddress` parameter received.');
        }
        try {
            const user: User = await this.graphClient
            .api(`/users/${emailAddress}`)
            .get();

            return user ? true : false;

        } catch (error) {
            return false;
        }
    }
}
