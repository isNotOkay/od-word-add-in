/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

const CURRENCY_PATTERN = "[$€£¥₹]";

let currentResultCount = 0;
let currentIndex = -1;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        const runButton = document.getElementById("run");
        const searchButton = document.getElementById("search-currency");
        const prevButton = document.getElementById("prev-result");
        const nextButton = document.getElementById("next-result");

        if (sideloadMsg && appBody) {
            sideloadMsg.style.display = "none";
            appBody.style.display = "flex";
        }

        if (runButton) {
            // Keep "Run" working: trigger the same search/highlight behavior
            runButton.onclick = () => {
                run().catch(console.error);
            };
        }

        if (searchButton) {
            searchButton.onclick = () => {
                searchCurrencies().catch(console.error);
            };
        }

        if (prevButton) {
            prevButton.onclick = () => {
                gotoPrevious().catch(console.error);
            };
        }

        if (nextButton) {
            nextButton.onclick = () => {
                gotoNext().catch(console.error);
            };
        }
    }
});

export async function run() {
    // For simplicity, make "Run" do the currency search + highlight as well.
    await searchCurrencies();
}

async function searchCurrencies() {
    await Word.run(async (context) => {
        const body = context.document.body;

        // Search for any of the currency symbols using a wildcard set.
        const results = body.search(CURRENCY_PATTERN, {
            matchCase: false,
            matchWholeWord: false,
            matchWildcards: true,
        });

        results.load("items");
        await context.sync();

        // Highlight all matches.
        results.items.forEach((r) => {
            r.font.highlightColor = "yellow";
        });

        currentResultCount = results.items.length;
        currentIndex = currentResultCount > 0 ? 0 : -1;

        // If we have at least one result, select the first one.
        if (currentIndex >= 0) {
            results.items[currentIndex].select();
        }

        await context.sync();
    });

    updateUI();
}

async function gotoNext() {
    if (currentResultCount === 0) {
        return;
    }

    currentIndex = (currentIndex + 1) % currentResultCount;
    updateUI();
    await selectCurrent();
}

async function gotoPrevious() {
    if (currentResultCount === 0) {
        return;
    }

    currentIndex = (currentIndex - 1 + currentResultCount) % currentResultCount;
    updateUI();
    await selectCurrent();
}

// Re-run the search and select the current index.
// We do NOT store Word.Range objects globally (they are context-bound).
async function selectCurrent() {
    if (currentIndex < 0 || currentResultCount === 0) {
        return;
    }

    await Word.run(async (context) => {
        const body = context.document.body;

        const results = body.search(CURRENCY_PATTERN, {
            matchCase: false,
            matchWholeWord: false,
            matchWildcards: true,
        });

        results.load("items");
        await context.sync();

        const actualCount = results.items.length;
        if (actualCount === 0) {
            return;
        }

        // Clamp index if document changed and number of results differs.
        const index = Math.min(currentIndex, actualCount - 1);
        results.items[index].select();

        await context.sync();
    });
}

function updateUI() {
    const infoEl = document.getElementById("result-info");
    const prevBtn = document.getElementById("prev-result") as HTMLButtonElement | null;
    const nextBtn = document.getElementById("next-result") as HTMLButtonElement | null;

    if (!infoEl) {
        return;
    }

    if (currentResultCount === 0) {
        infoEl.textContent = "No currency symbols found.";
        if (prevBtn) prevBtn.disabled = true;
        if (nextBtn) nextBtn.disabled = true;
        return;
    }

    infoEl.textContent = `Result ${currentIndex + 1} of ${currentResultCount}`;

    const disableNav = currentResultCount <= 1;
    if (prevBtn) prevBtn.disabled = disableNav;
    if (nextBtn) nextBtn.disabled = disableNav;
}
