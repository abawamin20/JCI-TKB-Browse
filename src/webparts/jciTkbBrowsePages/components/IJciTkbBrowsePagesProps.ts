import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IJciTkbBrowsePagesProps {
  context: WebPartContext;
  selectedViewId: string;
  feedbackPageUrl: string;
}
