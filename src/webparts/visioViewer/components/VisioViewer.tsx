import * as React from 'react';
import styles from './VisioViewer.module.scss';
import { IVisioViewerProps } from './IVisioViewerProps';
import { escape } from '@microsoft/sp-lodash-subset';
import 'officejs';

export default class VisioViewer extends React.Component<IVisioViewerProps, {}> {

  private _session: OfficeExtension.EmbeddedSession = null;


  public async componentDidMount() {
    let url = "https://zalodev.sharepoint.com/:u:/r/sites/VisioSite/_layouts/15/Doc.aspx?sourcedoc=%7B9354fede-6fc7-4f6d-90dd-e68499d364b2%7D&action=embedview&uid=%7B9354FEDE-6FC7-4F6D-90DD-E68499D364B2%7D&ListItemId=1&ListId=%7BD3496B90-A410-4841-A88B-44D2F247B30F%7D&odsp=1&env=prod&cid=b4f8b561-3c24-4ea4-9dc9-4ddae1869693";


    this._session = new OfficeExtension.EmbeddedSession(url,
      {
        id: styles.iframeHost,
        container: document.getElementById("iframeHost")
        // width: width,
        // height: height,
      });

    await this._session.init();
    this.visioLoaded();
  }

  private visioLoaded() {
    Visio.run(this._session, (ctx) => {
      let application = ctx.document.application;
      let doc = ctx.document;
      doc.view.hideDiagramBoundary = true;
      doc.view.disablePan = true;
      doc.view.disableZoom = true;
      application.showBorders = false;
      application.showToolbars = false;
      doc.onSelectionChanged.add(this.shapeClicked);

      return ctx.sync();
    });
  }

  private shapeClicked = async (args: Visio.SelectionChangedEventArgs): Promise<any> => {
    Visio.run(this._session, async (ctx) => {

      var page = ctx.document.getActivePage();
      var shapes = page.shapes;
      let shape = shapes.getItem(args.shapeNames[0]);
      shape.load();
      shape.view.load();
      let items = shape.shapeDataItems.load();
      await ctx.sync();
     shape.view.highlight = { color: "#41f444", width: 2 };
      if (items.items[0]) {
        alert((items.items[0].value));
      }
    });
  }

  private changePage = () => {
    Visio.run(this._session, async (ctx) => {
      ctx.document.setActivePage("Idea Phase");
      ctx.sync();
    });

  }


  public render(): React.ReactElement<IVisioViewerProps> {
    return (
      <div>
        <div id="iframeHost"></div>
        <a href="#" onClick={this.changePage}>Change the page</a>
      </div>
    );
  }
}
