// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { ActionContext } from "@typeagent/agent-sdk";
import { BrowserActionContext } from "../actionHandler.mjs";
import { BrowserConnector } from "../browserConnector.mjs";
import { createInstacartPageTranslator } from "./translator.mjs";
import {
  PurchaseResults,
  PurchaseSummary,
} from "../commerce/schema/shoppingResults.mjs";
import {
  AllListsInfo,
  AllRecipeSearchResults,
  BuyItAgainHeaderSection,
  BuyItAgainNavigationLink,
  HomeLink,
  ListDetailsInfo,
  ListInfo,
  ListsNavigationLink,
  NearbyStoresList,
  ProductDetailsHeroTile,
  ProductTile,
  RecipeHeroSection,
  SearchInput,
  StoreInfo,
} from "./schema/pageComponents.mjs";

export async function handleInstacartUserAction(
  action: any,
  context: ActionContext<BrowserActionContext>,
) {
  let message = "OK";
  if (!context.sessionContext.agentContext.browserConnector) {
    throw new Error("No connection to browser session.");
  }

  const browser: BrowserConnector =
    context.sessionContext.agentContext.browserConnector;

  const agent = await createInstacartPageTranslator("GPT_4_O_MINI");

  switch (action.actionName) {
    case "searchForProductAction":
      await handleFindProduct(action);
      break;
    case "addToCartAction":
      await handleAddToCart(action);
    case "addToListAction":
      await handleAddToList(action);
    case "findNearbyStoreAction":
      await handleFindStores(action);
      break;
    case "searchForRecipeAction":
      await handleFindRecipe(action);
      break;
    case "buyAllInRecipeAction":
      await handleBuyRecipeIngredients(action);
      break;
    case "buyAllInListAction":
      await handleBuyListContents(action);
      break;
    case "setPreferredStoreAction":
      await handleSetPreferredStore(action);
      break;
    case "buyItAgainAction":
      await handleBuyItAgain(action);
      break;
  }

  type UIElementSchemas = {
    AllListsInfo: AllListsInfo;
    AllRecipeSearchResults: AllRecipeSearchResults;
    BuyItAgainHeaderSection: BuyItAgainHeaderSection;
    BuyItAgainNavigationLink: BuyItAgainNavigationLink;
    HomeLink: HomeLink;
    ListDetailsInfo: ListDetailsInfo;
    ListInfo: ListInfo;
    ListsNavigationLink: ListsNavigationLink;
    NearbyStoresList: NearbyStoresList;
    ProductDetailsHeroTile: ProductDetailsHeroTile;
    ProductTile: ProductTile;
    RecipeHeroSection: RecipeHeroSection;
    SearchInput: SearchInput;
    StoreInfo: StoreInfo;
  };

  async function getPageComponent<T extends keyof UIElementSchemas>(
    componentType: T,
    selectionCondition?: string,
  ): Promise<UIElementSchemas[T] | undefined> {
    const htmlFragments = await browser.getHtmlFragments(true);
    const timerName = `getting ${componentType} section`;

    console.time(timerName);
    const response = await agent.getPageComponentSchema(
      componentType,
      selectionCondition,
      htmlFragments,
      undefined,
    );

    if (!response.success) {
      console.error("Attempt to get page component failed");
      console.error(response.message);
      return undefined;
    }

    console.timeEnd(timerName);
    return response.data as UIElementSchemas[T];
  }

  async function followLink(linkSelector: string | undefined) {
    if (!linkSelector) return;

    await browser.clickOn(linkSelector);
    await browser.awaitPageInteraction();
    await browser.awaitPageLoad();
  }

  async function searchOnWebsite<T extends keyof UIElementSchemas>(
    componentType: T,
    keywords: string,
  ): Promise<UIElementSchemas[T] | undefined> {
    if (componentType == "StoreInfo") {
      await goToHomepage();
    }

    const selector = await getPageComponent("SearchInput");
    if (!selector) {
      return;
    }

    const searchSelector = selector.cssSelector;

    await browser.clickOn(searchSelector);

    let queryPrefix = "";
    switch (componentType) {
      case "StoreInfo": {
        queryPrefix = "store: ";
        break;
      }
      case "AllRecipeSearchResults": {
        queryPrefix = "recipe: ";
        break;
      }
    }

    await browser.enterTextIn(queryPrefix + keywords, searchSelector);
    await browser.clickOn(selector.submitButtonCssSelector);
    await browser.awaitPageInteraction();
    await browser.awaitPageLoad();

    const request = `Search result: ${keywords}`;
    const result = await getPageComponent(componentType, request);

    return result as UIElementSchemas[T];
  }

  async function handleFindProduct(action: any) {
    const targetProduct = await searchOnWebsite(
      "ProductTile",
      action.parameters.keyword,
    );
    await followLink(targetProduct?.detailsLinkSelector);
  }

  async function handleAddToCart(action: any) {
    const targetProduct = await getPageComponent("ProductDetailsHeroTile");

    if (targetProduct?.addToCartButton) {
      await browser.clickOn(targetProduct.addToCartButton.cssSelector);
    }
  }

  async function handleAddToList(action: any) {
    const targetProduct = await getPageComponent("ProductDetailsHeroTile");

    if (targetProduct?.addToListButton) {
      await browser.clickOn(targetProduct.addToListButton.cssSelector);

      // this launches a popup with the available lists
      const request = `ListName: ${action.listName}`;
      const targetList = await getPageComponent("AllListsInfo", request);

      if (targetList?.lists) {
        await browser.clickOn(targetList.lists[0].cssSelector);
        await browser.clickOn(targetList.submitButtonCssSelector);
      }
    }
  }

  async function goToHomepage() {
    const link = await getPageComponent("HomeLink");
    console.log(link);

    await followLink(link?.linkCssSelector);
  }

  async function handleFindStores(action: any) {
    await goToHomepage();
    const storesList = await getPageComponent("NearbyStoresList");
    console.log(storesList);
    return storesList;
  }

  async function handleSetPreferredStore(action: any) {
    const targetStore = await searchOnWebsite(
      "StoreInfo",
      action.parameters.storeName,
    );
    await followLink(targetStore?.storeLinkCssSelector);

    // TODO: persist preferrences
  }

  async function handleFindRecipe(action: any) {
    const allRecipes = await searchOnWebsite(
      "AllRecipeSearchResults",
      action.parameters.keyword,
    );

    if (allRecipes && allRecipes.recipes.length > 0) {
      await followLink(allRecipes.recipes[0].recipeLinkCssSelector);
    }
  }

  async function handleBuyRecipeIngredients(action: any) {
    const allRecipes = await searchOnWebsite(
      "AllRecipeSearchResults",
      action.parameters.recipeName,
    );

    if (allRecipes && allRecipes.recipes.length > 0) {
      await followLink(allRecipes.recipes[0].recipeLinkCssSelector);

      const targetRecipe = await getPageComponent("RecipeHeroSection");

      if (targetRecipe?.addAllIngridientsCssSelector) {
        await browser.clickOn(targetRecipe.addAllIngridientsCssSelector);
      }
    }
  }

  async function handleBuyListContents(action: any) {
    await selectStore(action.parameters.storeName);

    const navigationLink = await getPageComponent("ListsNavigationLink");

    if (navigationLink?.linkCssSelector) {
      await followLink(navigationLink?.linkCssSelector);

      const request = `List name: ${action.parameters.listName}`;
      const targetList = await getPageComponent("ListInfo", request);

      if (targetList?.detailsLinkCssSelector) {
        await followLink(targetList.detailsLinkCssSelector);

        const listDetails = await getPageComponent("ListDetailsInfo");

        if (listDetails && listDetails.products) {
          for (let product of listDetails.products) {
            if (product.addToCartButtonCssSelector) {
              await browser.clickOn(product.addToCartButtonCssSelector);
            }
          }
        }
      }
    }
  }

  async function selectStore(storeName: string) {
    await goToHomepage();
    const request = `Store name: ${storeName}`;
    const targetStore = await getPageComponent("StoreInfo", request);

    console.log(targetStore);
    await followLink(targetStore?.storeLinkCssSelector);
  }

  async function handleBuyItAgain(action: any) {
    let results: PurchaseResults = {
      addedToCart: [],
      unavailable: [],
      storeName: action.parameters.storeName,
      deliveryInformation: "",
    };

    await selectStore(action.parameters.storeName);

    const navigationLink = await getPageComponent("BuyItAgainNavigationLink");

    console.log(navigationLink);

    if (navigationLink) {
      await followLink(navigationLink.linkCssSelector);

      const headerSection = await getPageComponent("BuyItAgainHeaderSection");
      console.log(headerSection);

      if (headerSection?.products) {
        if (action.parameters.allItems) {
          for (let product of headerSection.products) {
            if (product.availability == "Out of stock") {
              results.unavailable.push(product);
            } else {
              if (product.addToCartButtonCssSelector) {
                await browser.clickOn(product.addToCartButtonCssSelector);
                await browser.awaitPageInteraction();
                results.addedToCart.push(product);
              } else {
                results.unavailable.push(product);
              }
            }
          }
        } else {
          const request = `Product: ${action.productName}`;
          const targetProduct = await getPageComponent("ProductTile", request);
          if (targetProduct && targetProduct.addToCartButtonCssSelector) {
            await browser.clickOn(targetProduct.addToCartButtonCssSelector);
            await browser.awaitPageInteraction();
          }
        }
      }

      const friendlyMessage = await agent.getFriendlyPurchaseSummary(results);

      if (friendlyMessage.success) {
        message = (friendlyMessage.data as PurchaseSummary).formattedMessage;
      }
    }
  }

  return message;
}
