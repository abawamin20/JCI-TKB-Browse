import * as React from "react";
import type { IJciTkbBrowsePagesProps } from "./IJciTkbBrowsePagesProps";
import PagesList from "./PagesList/PagesList";
import { SPComponentLoader } from "@microsoft/sp-loader";

interface IJciTkbBrowsePagesState {}

export default class JciTkbBrowsePages extends React.Component<
  IJciTkbBrowsePagesProps,
  IJciTkbBrowsePagesState
> {
  constructor(props: IJciTkbBrowsePagesProps) {
    super(props);

    // Load CSS files
    const cssURLs = [
      "https://cdn.jsdelivr.net/npm/bootstrap@5.0.2/dist/css/bootstrap.min.css",
      "https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css",
    ];
    cssURLs.forEach((url) => SPComponentLoader.loadCss(url));
  }

  public render(): React.ReactElement<IJciTkbBrowsePagesProps> {
    return (
      <React.Fragment>
        <PagesList
          context={this.props.context}
          selectedViewId={this.props.selectedViewId}
          feedbackPageUrl={this.props.feedbackPageUrl}
        />
      </React.Fragment>
    );
  }
}
