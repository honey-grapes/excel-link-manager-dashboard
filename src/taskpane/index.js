import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { ToastContainer } from "react-toastify";
//Inline CSS and style loaders are used since CSS cannot be loaded from the HTML file
//This error also occurred in the default Excel Add-In welcome page
import "style-loader!css-loader!./taskpane.css";
import "style-loader!css-loader!react-toastify/dist/ReactToastify.css";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Link Manager Dashboard";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <div>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        <ToastContainer autoClose={2000} closeOnClick hideProgressBar={true} />
      </div>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

/* Initial render showing a progress bar */
render(App);

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
