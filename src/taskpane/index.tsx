import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { Taskpane } from "./components/Taskpane";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <Taskpane />
    </FluentProvider>
  );
});

/* Hot Module Reloading (Optional for Dev) */
if ((module as any).hot) {
  (module as any).hot.accept("./components/Taskpane", () => {
    const NextTaskpane = require("./components/Taskpane").Taskpane;
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <NextTaskpane />
      </FluentProvider>
    );
  });
}
