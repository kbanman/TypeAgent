// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { SearchHit } from "@microsoft/microsoft-graph-types";
import registerDebug from "debug";
import { GraphClient } from "./graphClient.js";

export class SearchClient extends GraphClient {
    private readonly logger = registerDebug("typeagent:graphUtils:searchClient");
    public constructor() {
        super("@search login");
    }

    public async findFiles(query: string): Promise<SearchHit | undefined> {
        const client = await this.ensureClient();
        this.logger(`searching for ${query}`);
        return client
            .api("/search/query")
            .post({
                requests: [
                    {
                        entityTypes: ["driveItem", "listItem", "list"],
                        query: {
                            queryString: query,
                        },
                    },
                ],
            });
    }
}

export async function createSearchGraphClient(): Promise<SearchClient> {
    return new SearchClient();
}
