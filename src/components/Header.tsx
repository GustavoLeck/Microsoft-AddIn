"use client";

import { Image, Title1, makeStyles, tokens } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  /** Logo for the light theme. */
  logo: string;
  /** Logo for the dark theme, swapped in by Tailwind's `dark:` variant. */
  logoDark: string;
  message: string;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "12px",
    padding: "40px 20px 24px",
    backgroundColor: tokens.colorNeutralBackground3,
  },
  message: {
    fontWeight: tokens.fontWeightRegular,
  },
});

export default function Header({ title, logo, logoDark, message }: HeaderProps) {
  const styles = useStyles();

  return (
    <header className={styles.header}>
      <Image
        src={logo}
        alt={title}
        title={title}
        className="h-auto w-56 max-w-full dark:hidden"
      />
      <Image
        src={logoDark}
        alt={title}
        title={title}
        className="hidden h-auto w-56 max-w-full dark:block"
      />
      <Title1 as="h1" className={styles.message}>
        {message}
      </Title1>
    </header>
  );
}
