"use client";

import {
  FluentProvider,
  makeStyles,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { useDarkMode, useOffice } from "@/lib/office";
import Header from "@/components/Header";

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

export default function TaskpanePage() {
  const office = useOffice();
  const dark = useDarkMode(office.status === "ready" && office.host != null);
  const styles = useStyles();

  return (
    <FluentProvider
      theme={dark ? webDarkTheme : webLightTheme}
      className={styles.root}
    >
      <Header
        logo="/assets/BeInc-Logo-light.png"
        logoDark="/assets/BeInc-Logo-dark.png"
        title="BeInc"
        message="Welcome"
      />

      <main className={styles.main}>test</main>
    </FluentProvider>
  );
}
