"use client";

import {
  FluentProvider,
  MessageBar,
  MessageBarBody,
  Spinner,
  makeStyles,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { DesignIdeas24Regular, LockOpen24Regular, Ribbon24Regular } from "@fluentui/react-icons";
import { useDarkMode, useOffice } from "@/lib/office";
import Header from "./Header";
import HeroList, { type HeroListItem } from "./HeroList";
import SubjectReader from "./SubjectReader";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
  },
  main: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
    padding: "10px 20px 24px",
  },
});

const features: HeroListItem[] = [
  { icon: <Ribbon24Regular />, primaryText: "Achieve more with Office integration" },
  { icon: <LockOpen24Regular />, primaryText: "Unlock features and functionality" },
  { icon: <DesignIdeas24Regular />, primaryText: "Create and visualize like a pro" },
];

export default function TaskpaneApp() {
  const office = useOffice();
  const dark = useDarkMode(office.status === "ready");
  const styles = useStyles();

  const inOutlook = office.status === "ready" && office.host === Office.HostType.Outlook;

  return (
    <FluentProvider theme={dark ? webDarkTheme : webLightTheme} className={styles.root}>
      <Header logo="/assets/logo-filled.png" title="BeInc" message="Welcome" />

      <main className={styles.main}>
        {office.status === "loading" && <Spinner label="Connecting to Outlook…" />}

        {office.status === "unavailable" && (
          <MessageBar intent="error">
            <MessageBarBody>Office.js could not be loaded. Check your internet connection.</MessageBarBody>
          </MessageBar>
        )}

        {office.status === "ready" && !inOutlook && (
          <MessageBar intent="warning">
            <MessageBarBody>
              Open this pane from Outlook. Sideload the add-in (manifest.xml) to see the app body.
            </MessageBarBody>
          </MessageBar>
        )}

        {inOutlook && (
          <>
            <HeroList message="Discover what BeInc can do for you today!" items={features} />
            <SubjectReader />
          </>
        )}
      </main>
    </FluentProvider>
  );
}
