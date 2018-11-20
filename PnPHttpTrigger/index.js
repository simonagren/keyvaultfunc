
const sp = require("@pnp/sp").sp;
const SPFetchClient = require("@pnp/nodejs").SPFetchClient;

const graph = require("@pnp/graph").graph;
const AdalFetchClient = require("@pnp/nodejs").AdalFetchClient;

const KeyVault = require('azure-keyvault');
const msRestAzure = require('ms-rest-azure');

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    if (req.query.site || (req.body && req.body.site)) {
        try {
            const siteName = req.query.site || req.body.site;

            const vaultUri = "https://simonsfuncvault.vault.azure.net/";
            // Should always be https://vault.azure.net
            const credentials = await msRestAzure.loginWithAppServiceMSI({resource: 'https://vault.azure.net'});

            const keyVaultClient = new KeyVault.KeyVaultClient(credentials);

            // Get SharePoint key value
            const spVaultSecret = await keyVaultClient.getSecret(vaultUri, "spSecret", "");
            const spSecret = spVaultSecret.value;

            // Get Graph key value
            const graphVaultSecret = await keyVaultClient.getSecret(vaultUri, "graphSecret", "");
            const graphSecret = graphVaultSecret.value;

            // Setup PnPJs via sp, with spSecret from Key Vault
            sp.setup({
                sp: {
                    fetchClientFactory: () => {
                        return new SPFetchClient(
                            `${process.env.spTenantUrl}/sites/${siteName}/`,
                            process.env.spId,
                            spSecret
                        );
                    },
                },
            });

            // Setup PnPJs Graph with graphSecret from Key Vault
            graph.setup({
                graph: {
                    fetchClientFactory: () => {
                        return new AdalFetchClient(
                            process.env.graphTenant,
                            process.env.graphId,
                            graphSecret
                        );
                    },
                },
            });

            // Get the web and select only Title
            const web = await sp.web.select("Title").get();
            
            // Filter Office365 groups for any with the same displayName as the web title
            const filtGroups = await graph.groups.filter(`displayName eq '${web.Title}'`).get();
            
            let createdTeam;
            // If only one group. It should only be one
            if (filtGroups.length === 1) {
                // Get the group from the array
                const group = filtGroups[0];
                // Create a Team based on that Group
                createdTeam = await graph.groups.getById(group.id).createTeam({
                    "memberSettings": {
                        "allowCreateUpdateChannels": true
                    },
                    "messagingSettings": {
                        "allowUserEditMessages": true,
                        "allowUserDeleteMessages": true
                    },
                    "funSettings": {
                        "allowGiphy": true,
                        "giphyContentRating": "strict"
                    }
                });
            }
            context.res = {
                // status: 200, /* Defaults to 200 */
                body: `Created a Team for the site ${web.Title}. Result: ${JSON.stringify(createdTeam)}`
            };
        } catch (error) {
            context.res = {
                status: 400,
                body: `Something went wrong: ${error}`
            }
        }
    }
    else {
        context.res = {
            status: 400,
            body: "Please pass a site on the query string or in the request body"
        };
    }

};