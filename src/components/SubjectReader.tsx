"use client";

import { useState } from "react";
import {
  Body1,
  Button,
  MessageBar,
  MessageBarBody,
  Text,
  makeStyles,
} from "@fluentui/react-components";
import { getItemSubject } from "@/lib/mailbox";

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
    marginTop: "8px",
    width: "100%",
  },
  result: {
    wordBreak: "break-word",
    textAlign: "center",
  },
});

/** Sample action ported from the Yeoman template: reads the subject of the email being written. */
export default function SubjectReader() {
  const styles = useStyles();
  const [subject, setSubject] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function run() {
    setBusy(true);
    setError(null);
    try {
      setSubject(await getItemSubject());
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setBusy(false);
    }
  }

  return (
    <section className={styles.section}>
      <Body1 align="center">
        Modify the source files, then click <b>Run</b>.
      </Body1>

      <Button appearance="primary" size="large" onClick={run} disabled={busy}>
        Run
      </Button>

      {subject !== null && (
        <Text className={styles.result}>
          <b>Subject:</b>
          <br />
          {subject || "(no subject)"}
        </Text>
      )}

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}
    </section>
  );
}
