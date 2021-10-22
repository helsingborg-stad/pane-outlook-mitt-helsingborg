import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./components/Header";
import HeroList, { HeroListItem } from "./components/HeroList";
import Progress from "./components/Progress";

// images references in the manifest
import "../assets/icon-16.png";
import "../assets/icon-32.png";
import "../assets/icon-80.png";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default (props: AppProps) => {
  const { title, isOfficeInitialized } = props;

  const listItems = [
    {
      icon: "Ribbon",
      primaryText: "Achieve more with Office integration",
    },
    {
      icon: "Unlock",
      primaryText: "Unlock features and functionality",
    },
    {
      icon: "Design",
      primaryText: "Create and visualize like a pro",
    },
  ];

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <div className="ms-welcome">
      <Header logo={require("../hbg-github-logo-combo.png")} title={title} message="Welcome" />
      <HeroList message="Discover what Office Add-ins can do for you today!" items={listItems}>
        <p className="ms-font-l">
          Modify the source files, then click <b>Run</b>.
        </p>
        <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }}>
          Run
        </DefaultButton>
      </HeroList>
    </div>
  );
};
