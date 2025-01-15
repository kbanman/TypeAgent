// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { createSearchGraphClient, SearchClient } from "graph-utils";
import chalk from "chalk";
import {
    SearchAction,
} from "./searchActionsSchema.js";
import { SearchHit } from "@microsoft/microsoft-graph-types";
import { ActionContext, AppAgent, SessionContext } from "@typeagent/agent-sdk";
import { createActionResultFromHtmlDisplay } from "@typeagent/agent-sdk/helpers/action";
import {
    CommandHandlerNoParams,
    CommandHandlerTable,
    getCommandInterface,
} from "@typeagent/agent-sdk/helpers/command";
import {
    displayStatus,
    displaySuccess,
    displayWarn,
} from "@typeagent/agent-sdk/helpers/display";

import registerDebug from "debug";
const debug = registerDebug("typeagent:search");

class SearchClientLoginCommandHandler implements CommandHandlerNoParams {
    public readonly description = "Log into the MS Graph to access search";
    public async run(context: ActionContext<SearchActionContext>) {
        const { searchClient } = context.sessionContext.agentContext;
        if (searchClient === undefined) {
            throw new Error("Search client not initialized");
        }
        if (searchClient.isAuthenticated()) {
            displayWarn("Already logged in", context);
            return;
        }

        await searchClient.login((prompt) => displayStatus(prompt, context));
        displaySuccess("Successfully logged in", context);
    }
}

class SearchClientLogoutCommandHandler implements CommandHandlerNoParams {
    public readonly description = "Log out of MS Graph to access search";
    public async run(context: ActionContext<SearchActionContext>) {
        const { searchClient } = context.sessionContext.agentContext;
        if (searchClient === undefined) {
            throw new Error("Search client not initialized");
        }
        if (searchClient.logout()) {
            displaySuccess("Successfully logged out", context);
        } else {
            displayWarn("Already logged out", context);
        }
    }
}

const handlers: CommandHandlerTable = {
    description: "Search login command",
    defaultSubCommand: "login",
    commands: {
        login: new SearchClientLoginCommandHandler(),
        logout: new SearchClientLogoutCommandHandler(),
    },
};

export function instantiate(): AppAgent {
    return {
        initializeAgentContext: initializeSearchContext,
        updateAgentContext: updateSearchContext,
        executeAction: executeSearchAction,
        ...getCommandInterface(handlers),
    };
}

type SearchActionContext = {
    searchClient: SearchClient | undefined;
};

async function initializeSearchContext() {
    return {
        searchClient: undefined,
    };
}

async function updateSearchContext(
    enable: boolean,
    context: SessionContext<SearchActionContext>,
): Promise<void> {
    if (enable) {
        context.agentContext.searchClient = await createSearchGraphClient();
    } else {
        context.agentContext.searchClient = undefined;
    }
}

async function executeSearchAction(
    action: SearchAction,
    context: ActionContext<SearchActionContext>,
) {
    const { searchClient } = context.sessionContext.agentContext;
    if (searchClient === undefined) {
        throw new Error("Search client not initialized");
    }

    if (!searchClient.isAuthenticated()) {
        await searchClient.login();
    }

    let result = await handleSearchAction(action, context);
    if (result) {
        return createActionResultFromHtmlDisplay(result);
    }
}

async function handleSearchAction(
    action: SearchAction,
    context: ActionContext<SearchActionContext>,
) {
    const { searchClient } = context.sessionContext.agentContext;
    if (!searchClient) {
        return "<div>Search client not initialized ...</div>";
    }

    let res;
    switch (action.actionName) {
        case "findFiles":
            debug(chalk.green("Handling findFiles action ..."));
            res = await searchClient.findFiles(action.parameters.query);
            return findResultsDisplayHtml(res);
            break;
        default:
            throw new Error(`Unknown action: ${(action as any).actionName}`);
    }
}

function findResultsDisplayHtml(hits: SearchHit[]): string {
    console.log('hits', hits);
    if (!hits || hits.length === 0) {
        return "";
    }
    let htmlEvents: string = `<div style="height: 100px; overflow-y: scroll; border: 1px solid #ccc; padding: 10px; font-family: sans-serif; font-size: small">SharePoint Search Results`;
    hits.forEach((hit) => {
        const link = hit.resource?.id;
        htmlEvents +=
            `<p><a href="${link}" target="_blank">` +
            `${hit.summary}</a>` +
            `</p>`;
    });
    htmlEvents += `</div>`;
    return htmlEvents;
}
