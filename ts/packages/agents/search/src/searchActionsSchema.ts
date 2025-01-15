// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

export type SearchAction =
    | FindFilesAction;

// Type for sending a simple email
export type FindFilesAction = {
    actionName: "findFiles";
    parameters: {
        // Subject of the email, infer the subject based on the user input
        query: string;
    };
};
