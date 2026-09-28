"use client";

import { useEffect, useState } from "react";
import {
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
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    function loop() {
      console.log("loop tick", new Date().toISOString());
    }
    const id = setInterval(loop, 1000);
    return () => clearInterval(id);
  }, []);

  return <section className={styles.section}></section>;
}
