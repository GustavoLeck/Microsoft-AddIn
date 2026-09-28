"use client";

import type { ReactElement } from "react";
import { Body1, Subtitle1, makeStyles, tokens } from "@fluentui/react-components";

export interface HeroListItem {
  icon: ReactElement;
  primaryText: string;
}

export interface HeroListProps {
  message: string;
  items: HeroListItem[];
}

const useStyles = makeStyles({
  list: {
    listStyleType: "none",
    margin: 0,
    padding: 0,
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  item: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
  },
  icon: {
    display: "flex",
    color: tokens.colorBrandForeground1,
  },
});

export default function HeroList({ message, items }: HeroListProps) {
  const styles = useStyles();

  return (
    <>
      <Subtitle1 as="h2" align="center">
        {message}
      </Subtitle1>
      <ul className={styles.list}>
        {items.map((item) => (
          <li className={styles.item} key={item.primaryText}>
            <span className={styles.icon}>{item.icon}</span>
            <Body1>{item.primaryText}</Body1>
          </li>
        ))}
      </ul>
    </>
  );
}
