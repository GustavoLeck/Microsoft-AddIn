import type { Metadata } from "next";
import TaskpaneLoader from "./TaskpaneLoader";

export const metadata: Metadata = {
  title: "BeInc",
};

export default function TaskpanePage() {
  return <TaskpaneLoader />;
}
