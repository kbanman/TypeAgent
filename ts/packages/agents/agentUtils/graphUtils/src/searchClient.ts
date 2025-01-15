// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { SearchHit, SearchResponse } from "@microsoft/microsoft-graph-types";
import registerDebug from "debug";
import { GraphClient } from "./graphClient.js";

export class SearchClient extends GraphClient {
    private readonly logger = registerDebug("typeagent:graphUtils:searchClient");
    public constructor() {
        super("@search login");
    }

    public async findFiles(query: string): Promise<SearchHit[]> {
        const client = await this.ensureClient();
        this.logger(`searching for ${query}`);
        const result = await client
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
            }) as SearchResponse;
        if (!result || !result.hitsContainers) {
            return [];
        }
        return result.hitsContainers.map((container) => container.hits).flat().filter<SearchHit>(isSearchHit);
    }
}

function isSearchHit(value: SearchHit | null | undefined): value is SearchHit {
    return value !== null && value !== undefined;
}

export async function createSearchGraphClient(): Promise<SearchClient> {
    return new SearchClient();
}
