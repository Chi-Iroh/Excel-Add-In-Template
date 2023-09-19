/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { Cell } from "../cell/Cell";
import { Direction } from "../cell/CellLocation";
import { handleError } from "../utils";
import { getWholeWorksheetRange } from "../excel/utils";

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("sideload-msg")!.style.display = "none";
        document.getElementById("app-body")!.style.display = "flex";
        document.getElementById("run")!.onclick = run;
    }
});

function fibonacci(context : Excel.RequestContext) {
    const currentWorksheet : Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();
    let first : Cell = new Cell(currentWorksheet, "A1");
    let second : Cell = new Cell(currentWorksheet, "A2");
    first.setValue(1);
    second.setValue(1);
    for (let i = 0; i < 25; i++) {
        const third = second.nextCell(1, Direction.Down);
        if (third == null) {
            return;
        }

        third.setValue(`=${first.getLocation()} + ${second.getLocation()}`);
        first = second;
        second = third;
    }
}
async function run() {
    await Excel.run(async (context : Excel.RequestContext) => {
        const currentWorksheet : Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();

        try {
            fibonacci(context);
        } catch (error) {
            console.error(error)
            handleError(error as Error);
        }

        const wholeRange = getWholeWorksheetRange(currentWorksheet);
        wholeRange.format.autoIndent = true;
        wholeRange.format.autofitColumns();
        wholeRange.format.autofitRows();

        await context.sync();
    });
}