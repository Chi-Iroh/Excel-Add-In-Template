# <a id="top"></a> Setup Office Add-In
## [Go to Microsoft auto-generated README](../README.md)

<br>

### Setup hook

To enable the commit message checking hook : `ln -T hooks/prepare-commit-msg .git/hooks/prepare-commit-msg`.  

## Introduction and Configuration

Office suite (Excel, Word etc..) supports add-ins, either in web or app mode (app is Windows-only).  
This document will explain how to create one, step by step.  
A dual boot (Windows 10 / Ubuntu 22.04 LTS) was used but steps shouldn't be OS-dependant.  


## Optional Step : Creating a partition for testing purposes ([skip](#after-partition))

In a dual boot context, it is interesting to make a partition which can be accessed from both Windows and another OS, so that the add-in can be tested easily on multiple platforms and in web or app mode.  

### Partition requirements

The filesystem must support multiple OSes :
* NTFS : recent, Windows and Linux only
* FAT32 : older, but supported almost anywhere (including MacOS)

Partition size hasn't to be very big, <ins>a few gigabytes</ins> seem to be more than enough.  

To shrink a partition and make space for this one, follow [this guide](https://access.redhat.com/articles/1196333) (Linux + e2fsck + resize2fs).  

A partition can be created either using a GUI (not covered here) or in the terminal (fdisk and mkfs are explained).  

### Creating the partition (Linux + fdisk)

Using fdisk, here's the steps (assuming there's some unallocated space in a disk) :  

| Prompt | Command | Explanation |
| :----- | :-----: | :---------: |
| None | `sudo fdisk /dev/XXX` | Opens the disk.  |
| `Command (m for help):` | `n` | Adds a new partition.  |
| `Partition number (...):` | `[Enter]` | Sets the partition number to Y (entering another number than the one proposed is OK, it will only affect device filename). Will create later a device block named `/dev/sdaY` or `/dev/nvmeAnB/Y` depending of disk type. |
| `First sector (...):` | `[Enter]` | Makes the partition begin right at the beginning of unallocated region in the disk. |
| `Last sector, +sectors or +size (...):` | `+XM` or `+XG` | Reserves X megabytes or gigabytes. |
| `Command (m for help):` | `t` | Partition type must be changed to be recognized by Windows. |
| `Partition number (...):` | `[Enter] or [partition number]` | Selects the partition number X. |
| `Partition type or alias (type L to list all):` | `11` (Microsoft basic data) | Partition will be recognized and automatically mounted on Windows. When typing L to see all the types, it opens a vim-like list which must be exited by entering q. |
| `Command (m for help):` | `w` | Applies the changes onto the disk. |

### Creating the filesystem (Linux + mkfs)

The partition is created but cannot be used until it has a filesystem :
* NTFS : `mkfs.ntfs [-L LABEL] [-f]` : -f is for quick format (full format is long)
* FAT32 : `mkfs.fat -F 32 [-n LABEL]`

The partition is now ready.  

## <a id="after-partition"></a> Generating Project Tree

Firstly, Node.js must be installed.
Then Yeoman will be used to generate the project, it must be installed alongside Office add-in generator like this : `npm install -g yo generator-office`.  
Run yeoman : `yo office` (no need to create the project directory by yourself, yeoman will do it).  
It will let you choose between several project types : choose <ins>Excel Custom Functions using a Shared Runtime </ins>.  
Then comes the language (JavaScript or TypeScript), the project name and the application supported (Word, Excel etc..).  

<img src="assets/Excel Add-in.png" alt="Custom add-in logo is on top right corner of Excel (in Home) -- Taskpane is on the right side.">

Note: <ins>Excel Custom Functions using a Shared Runtime</ins> provides both a taskpane and custom functions, whereas <ins>Office Add-in Task Pane project</ins> only provides a taskpane.  

## Configuring the compiler

```json
// tsconfig.json
{
    "compilerOptions": {
        "strict": true,     // More type-checking
        "forceConsistentCasingInFileNames": true,   // Some OSes (like Windows) treat lowercase and uppercase characters in the same way, this options forces to use the exact name and casing for compatibility.
        ...
    }
}
```

## [Excel JavaScript API Reference](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

## Running Unit Tests

`npm install --save-dev jest @types/jest ts-jest`  
* `jest` : Unit test framework (Javascript)  
* `ts-jest` : Typescript support  
* `@types/jest` or `@jest/globals` : Functions (`test`, `expect` etc..).   
  Note that you need to import them when using `@jest/globals`, but not with `@types/jest`.  
  Note that `@types/jest` may not be up-to-date (but still very recent), contrary to `@jest/globals` which is the latest version.

```json
// jest.config.json
{
    "preset": "ts-jest",        // Typescript support
    "testMatch": [ "**/*.ts" ], // Test all .ts files...
    "roots": ["tests/"],        // in "tests" directory
    "verbose": true             // Adds log for each tets
}
```

```json
// package.json
{
    ...
    "scripts": {
        ...
        "tests" : "jest",    // "npm run tests" to run jest.
        "tests" : "jest",    // "npm run tests" to run jest.
        ...
    }
}
```

[Jest Testing API](https://jestjs.io/docs/api)  

### Troubleshoot : `ReferenceError: Office / Excel is not defined`

This error may be triggered by Jest when you test a function which is implemented in a file which contains `Office` or `Excel` in the global namespace (not in a function nor a class, or in a function which is called in the global namespace).  

```typescript
// Yeoman's auto-generated src/taskpane/taskpane.ts
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        ...
    }
}
```

Jest (and any other unit test framework like `mocha`) can't find Office.js library because it's a single file in Microsoft's servers :  

```html
<!-- src/commands/commands.html -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
```

Microsoft then released a package named `office-addin-mock` (must be installed with `npm install --save-dev`) which will provide a mock for Office, Excel and other classes.  
A mock is a fake object which overrides the original one's behaviour.  
In this situation it can help a lot , we may make a mock of an Excel worksheet and give it a bidimensionnal array of numbers, and then test our logic on it without using a real file.  
Here's how is an unit test :  

```typescript
// tests/helloworld.ts
import { OfficeMockObject } from "office-addin-mock";

// Overriding methods according to the needs.
const officeMock = {
  workbook: {
    range: {
      address: "C2:G3",
    },
    getSelectedRange: function () {
      return this.range;
    },
  },
  onReady: async function () {}
};

const mock = new OfficeMockObject(officeMock);
global.Office = mock as any; // as any is forced because the compiler will complain
// since the mock hasn't the exact same methods and members as the original Office class (here we didn't implement everything).

// importing our test, after the mock so that Office will be found.
import { helloworld } from "../src/taskpane/taskpane";

describe("Test", () => {
    test("Hello, World", () => {                    // it("Hello, World", ...) also exists and behaves exactly as 'test'
      expect(helloworld()).toBe("Hello, World !");
    });
})
```

Note: If the tested function is in a file without Office, Excel or something like that in the global namespace, it's not needed to create the mock nor move the import at the end.  
Also if you need to reset the mock between two tests, it may be better to use `require(...)` inside the test :  

```typescript
import { OfficeMockObject } from "office-addin-mock";
const officeMock = ...;

describe("Test", () => {
    test("Hello, World", () => {
        global.Office = new OfficeMockObject(officeMock) as any;
        const helloworld = require("../src/taskpane/taskpane.ts");
        expect(helloworld.helloworld()).toBe("Hello, World !");
    }) // 'test' is mandatory and contains your assertions...
}) // but 'describe' is optional and just groups some tests or other 'describe's. Nested 'describe's are useful for more accurate messages.
```

Here the mock can be reinitialized or setup differently between tests.  
[GitHub PR on OfficeDev repo](https://github.com/OfficeDev/Office-Addin-TaskPane/pull/136/files/bbd173c3185d39cf8b3ef6364ebf8dcec62f7347)

## More on Unit Tests and Mocks

As said previously, a mock is useful since it emulates an object, in this case because Excel API isn't available for testing.  
So for testing purposes, a minimal implementation of Excel is provided in [tests/setup/setup.ts](../tests/setup/setup.ts).  
This mainly includes `Excel.Range` class, with a global bidimensionnal array acting as Excel worksheet.  
Constants `EXCEL_ROWS_MAX` and `EXCEL_COLUMNS_MAX` are redefined in this file to speed up tests, and thus can be changed again according to the needs.  
A minimal expression interpreter is provided, permitting to test expressions like `=A1+A2*A3+SQRT(A4)` (within the limits of what is implemented by Math.js).  
Here's how looks unit tests using [setup.ts](../tests/setup/setup.ts) mocks :  

```ts
// tests/someTest.ts
import "./setup/setup";
import { Cell } from "../src/cell/Cell"; // if needed, maybe accompanied with some other imports

test("someTest", () => {
  // Excel.RequestContext() simulates 'context' from src/taskpane/taskpane.ts
  const worksheet = new Excel.RequestContext().workbook.worksheets.getActiveWorksheet();
  const A1 = new Cell(worksheet, "A1");

  A1.setValue(1);
  A1.updateAllValues(); // explanation below
  expect(A1.getValue()).toBe(1);
  expect(new Cell(worksheet, "A1")).toBe(1); // the new instance retrieves the new value of A1
})
```

To simulate an Excel worksheet, there's a global bidimensionnal array (of either strings or numbers) in setup.ts.  
When `CellInstance.setValue(x)` is called, the internal Excel.Range object is updated, but only locally.  
To update the global array, one needs to call `CellInstance.updateAllValues()` if manipulating a `Cell` instance, or `ExcelRangeInstance.updateAllValues()` when directly using `Excel.Range` (`Cell` performs updating by calling `Excel.Range`'s mock method).  
Note: Forgetting to call `.updateAllValues()` before an assertion is very likely to make the test fail.  

## Requesting an API and dealing with CORS

Using a web API implies making requests to a (maybe external) server.  
For security reasons, most web browsers prevents a script (JavaScript or TypeScript backend) from requesting an external URL (an URL which is not in the same machine as the script).  
Whatever server it is, that means the owner didn't enable CORS, intentionally or not (if not, one should contact him).  
Useful resources :
* [Mozilla documentation](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS)
* [Setting up CORS on the most common servers](https://enable-cors.org/server.html)

### [Go back to top](#top)