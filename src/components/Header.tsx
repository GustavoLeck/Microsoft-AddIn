"use client";

import { Image, Title1, makeStyles, tokens } from "@fluentui/react-components";

export interface HeaderProps {
  title: string;
  logo: string;
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

export default function Header({ title, logo, message }: HeaderProps) {
  const styles = useStyles();

  return (
    <header className={styles.header}>
      <Image width={90} height={90} src={logo} alt={title} title={title} />
      <Title1 as="h1" className={styles.message}>
        {message}
      </Title1>
    </header>
  );
}
