/* global document, Office, module, require */
import App from "./src/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import React from "react";
import ReactDOM from "react-dom";

import GlobalStyle from "./globalStyle";

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <>
        <ThemeProvider>
          <GlobalStyle />
          <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </ThemeProvider>
      </>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

if ((module as any).hot) {
  (module as any).hot.accept("./src/App", () => {
    const NextApp = require("./src/App").default;
    render(NextApp);
  });
}
