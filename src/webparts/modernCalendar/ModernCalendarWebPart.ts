/* eslint-disable @typescript-eslint/no-explicit-any */
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneDropdown,

} from "@microsoft/sp-property-pane";

import { Version } from "@microsoft/sp-core-library";

import * as strings from "modernCalendarStrings";
import { IModernCalendarWebPartProps } from "./IModernCalendarWebPartProps";
import CalendarTemplate from "./CalendarTemplate";

import jQuery from "jquery";
import moment from 'moment';
//import Swal from "sweetalert2";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { EventDetailsPanel, IEventDetails } from './components/EventDetailsPanel';

import { Calendar, EventClickArg, EventSourceInput, EventMountArg } from '@fullcalendar/core';
import dayGridPlugin from '@fullcalendar/daygrid';
import listPlugin from '@fullcalendar/list'; // Assuming list view is used or supported, otherwise just dayGridPlugin if listWeek provided by dayGrid (it might be separate)
// Actually standard FullCalendar package usually needs interaction plugin for clicks? No, core handles it.
// Checking imports.




export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}



export default class ModernCalendarWebPart extends BaseClientSideWebPart<IModernCalendarWebPartProps> {
  public constructor() {
    super();
  }

  private _log(level: 'debug' | 'info' | 'warn' | 'error', message: string, data?: any): void {
    const payload: any = {
      ts: new Date().toISOString(),
      level,
      message,
    };
    if (data !== undefined) {
      payload.data = data;
    }
    // Prefix so it's easy to find in logs
    const prefix = '[ModernCalendarWebPart]';
    try {
      if (level === 'error') {
        console.error(prefix, payload);
      } else if (level === 'warn') {
        console.warn(prefix, payload);
      } else if (level === 'debug' && typeof console.debug === 'function') {
        console.debug(prefix, payload);
      } else {
        console.log(prefix, payload);
      }
    } catch {
      // Fallback to plain console.log if structured logging fails
      console.log(prefix, message, data);
    }
  }

  private _stringifyError(err: unknown): string {
    try {
      if (err instanceof Error) {
        const stack = (err as any).stack ? `\n${(err as any).stack}` : '';
        return `${err.name}: ${err.message}${stack}`;
      }
      if (typeof err === 'string') return err;
      if (err == null) return 'Unknown error (null/undefined)';
      return JSON.stringify(err, Object.getOwnPropertyNames(err));
    } catch {
      return String(err);
    }
  }

  private _renderFatalError(userMessage: string, err?: unknown): void {
    const details = err ? this._stringifyError(err) : '';
    if (err) console.error('[ModernCalendarWebPart] Fatal error:', err);
    const msg = details ? `${userMessage}\n\nDetails:\n${details}` : userMessage;
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.context.statusRenderer.renderError(this.domElement, msg);
  }

  public render(): void {
    if (this.properties.theme != null) {
      SPComponentLoader.loadCss(this.properties.theme);
    }

    if (!this.properties.other) {
      jQuery("input[aria-label=hide-col]").parent().hide();
    }

    //Check required properties before rendering list
    if (
      this.properties.listTitle == null ||
      this.properties.start == null ||
      this.properties.end == null ||
      this.properties.title == null ||
      this.properties.detail == null
    ) {
      const missing: string[] = [];
      if (!this.properties.listTitle) missing.push("listTitle");
      if (!this.properties.start) missing.push("start");
      if (!this.properties.end) missing.push("end");
      if (!this.properties.title) missing.push("title");
      if (!this.properties.detail) missing.push("detail");
      if (!this.properties.colorField) missing.push("colorField");
      this._log('warn', 'Missing required properties for ModernCalendarWebPart', { missing });
      this.domElement.innerHTML = CalendarTemplate.emptyHtml(this.properties.description);
    } else {
      this.domElement.innerHTML = CalendarTemplate.templateHtml;
      this._renderListAsync();
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected onPropertyPaneConfigurationStart(): void {
    //Set a default theme
    if (this.properties.theme == null) {
      this.properties.theme = CalendarTemplate.theme()[0].key.toString();
    }

    if (this.properties.site) {
      this.listDisabled = false;
    }

    if (
      this.properties.listTitle &&
      (!this.properties.start ||
        !this.properties.end ||
        !this.properties.title ||
        !this.properties.detail ||
        !this.properties.colorField)
    ) {
      this._log('debug', 'this.properties.listTitle', { listTitle: this.properties.listTitle });

      this._getListColumns(this.properties.listTitle, this.properties.site).then((response3) => {
        const col: IPropertyPaneDropdownOption[] = [];
        for (const _innerKey in response3.value) {
          col.push({
            key: response3.value[_innerKey]["InternalName"],
            text: response3.value[_innerKey]["Title"],
          });
        }
        this._columnOptions = col;
        this.colsDisabled = false;
        this.listDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(
          this.domElement
        );
        this.render();
      });
    }

    if (!this.properties.other) {
      jQuery("input[aria-label=hide-col]").parent().hide();
    }

    if (
      this.properties.site &&
      this.properties.listTitle &&
      this.properties.start &&
      this.properties.start &&
      this.properties.end &&
      this.properties.title &&
      this.properties.detail &&
      this.properties.colorField
    ) {
      this.context.statusRenderer.displayLoadingIndicator(
        this.domElement,
        "Configuration"
      );
      this._getSiteRootWeb().then((response0) => {
        this._getSites(response0["Url"]).then((response) => {
          const sites: IPropertyPaneDropdownOption[] = [];
          sites.push({
            key: this.context.pageContext.web.absoluteUrl,
            text: "This Site",
          });
          sites.push({ key: "other", text: "Other Site (Specify Url)" });
          for (const _key in response.value) {
            if (
              this.context.pageContext.web.absoluteUrl !=
              response.value[_key]["Url"]
            ) {
              sites.push({
                key: response.value[_key]["Url"],
                text: response.value[_key]["Title"],
              });
            }
          }
          this._siteOptions = sites;
          if (this.properties.site) {
            this._getListTitles(this.properties.site).then((response2) => {
              this._dropdownOptions = response2.value.map((list: ISPList) => {
                return {
                  key: list.Title,
                  text: list.Title,
                };
              });
              this._log('debug', 'this.properties.site', { site: this.properties.site });
              this._getListColumns(
                this.properties.listTitle!,
                this.properties.site
              ).then((response3) => {
                const col: IPropertyPaneDropdownOption[] = [];
                for (const _innerKey in response3.value) {
                  col.push({
                    key: response3.value[_innerKey]["InternalName"],
                    text: response3.value[_innerKey]["Title"],
                  });
                }
                this._columnOptions = col;
                this.colsDisabled = false;
                this.listDisabled = false;
                this.context.propertyPane.refresh();
                this.context.statusRenderer.clearLoadingIndicator(
                  this.domElement
                );
                this.render();
              });
            });
          }
        });
      });
    } else {
      this._getSitesAsync();
    }
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (newValue == "other") {
      this.properties.other = true;
      this.properties.listTitle = null;
      jQuery("input[aria-label=hide-col]").parent().show();
    } else if (oldValue === "other" && newValue != "other") {
      this.properties.other = false;
      this.properties.siteOther = null;
      this.properties.listTitle = null;
      jQuery("input[aria-label=hide-col]").parent().hide();
    }
    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "Configuration"
    );
    if ((propertyPath === "site" || propertyPath === "siteOther") && newValue) {
      this.colsDisabled = true;
      this.listDisabled = true;
      let siteUrl = newValue;
      if (this.properties.other && this.properties.siteOther) {
        siteUrl = this.properties.siteOther;
      } else {
        jQuery("input[aria-label=hide-col]").parent().hide();
      }
      if (
        (this.properties.other && this.properties.siteOther && this.properties.siteOther.length > 25) ||
        !this.properties.other
      ) {
        this._getListTitles(siteUrl).then((response) => {
          this._dropdownOptions = response.value.map((list: ISPList) => {
            return {
              key: list.Title,
              text: list.Title,
            };
          });

          this.listDisabled = false;
          this.context.propertyPane.refresh();
          this.context.statusRenderer.clearLoadingIndicator(this.domElement);
          this.render();
        }).catch(() => {
          this._log('error', 'Error loading lists');
        });
      }
    } else if (propertyPath === "listTitle" && newValue) {
      // tslint:disable-next-line:no-duplicate-variable
      let siteUrl = this.properties.site;
      if (this.properties.other && this.properties.siteOther) {
        siteUrl = this.properties.siteOther;
      }
      this._log('debug', 'siteUrl', { siteUrl });
      this._getListColumns(newValue, siteUrl).then((response) => {
        const col: IPropertyPaneDropdownOption[] = [];
        for (const _key in response.value) {
          col.push({
            key: response.value[_key]["InternalName"],
            text: response.value[_key]["Title"],
          });
        }
        this._columnOptions = col;
        this.colsDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });
    } else {
      //Handle other fields here
      this.render();
    }
  }

  private colsDisabled: boolean = true;
  private listDisabled: boolean = true;

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    let otherSiteAria = "hide-col";
    if (this.properties.other) {
      otherSiteAria = "";
    }
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
                PropertyPaneDropdown("theme", {
                  label: "Theme",
                  options: CalendarTemplate.theme(),
                }),
                PropertyPaneDropdown("site", {
                  label: "Site",
                  options: this._siteOptions,
                }),
                PropertyPaneTextField("siteOther", {
                  label:
                    "Other Site Url (i.e. https://contoso.sharepoint.com/path)",
                  ariaLabel: otherSiteAria,
                }),
                PropertyPaneDropdown("listTitle", {
                  label: "List Title",
                  options: this._dropdownOptions,
                  disabled: this.listDisabled,
                }),
                PropertyPaneDropdown("start", {
                  label: "Start Date Field",
                  options: this._columnOptions,
                  disabled: this.colsDisabled,
                }),
                PropertyPaneDropdown("end", {
                  label: "End Date Field",
                  options: this._columnOptions,
                  disabled: this.colsDisabled,
                }),
                PropertyPaneDropdown("title", {
                  label: "Event Title Field",
                  options: this._columnOptions,
                  disabled: this.colsDisabled,
                }),
                PropertyPaneDropdown("detail", {
                  label: "Event Details",
                  options: this._columnOptions,
                  disabled: this.colsDisabled,
                }),
                PropertyPaneDropdown("colorField", {
                  label: "Event Color Field (Hex/RGB)",
                  options: this._columnOptions,
                  disabled: this.colsDisabled,
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private _siteOptions: IPropertyPaneDropdownOption[] = [];
  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];
  private _columnOptions: IPropertyPaneDropdownOption[] = [];



  private _getSiteRootWeb(): Promise<any> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
        `/_api/Site/RootWeb?$select=Title,Url`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getSites(rootWebUrl: string): Promise<any> {
    return this.context.spHttpClient
      .get(
        rootWebUrl + `/_api/web/webs?$select=Title,Url`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getListTitles(site: string): Promise<any> {
    return this.context.spHttpClient
      .get(
        site + `/_api/web/lists?$filter=Hidden eq false and BaseType eq 0`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        this._log('debug', 'response get List Titles');
        return response.json();
      });
  }

  private _getListColumns(
    listNameColumns: string,
    listsite: string
  ): Promise<any> {
    return this.context.spHttpClient
      .get(
        listsite +
        `/_api/web/lists/GetByTitle('${listNameColumns}')/Fields?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        this._log('debug', 'listsite', { listsite });
        return response.json();
      });
  }

  private _getListData(listName: string, site: string): Promise<any> {
    this._log('debug', 'listName', { listName });
    return this.context.spHttpClient
      .get(
        site +
        `/_api/web/lists/GetByTitle('${listName}')/items?$select=${encodeURIComponent(
          this.properties.title
        )},${encodeURIComponent(this.properties.start)},${encodeURIComponent(
          this.properties.end
        )},${encodeURIComponent(
          this.properties.detail
        )},${encodeURIComponent(
          this.properties.colorField
        )},Created,Author/ID,Author/Title&$expand=Author/ID,Author/Title&$orderby=Id desc&$limit=500`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        this._log('debug', 'response get List Data', { response });
        return response.json();
      });
  }

  private _renderList(items: any[]): void {
    const calItems: EventSourceInput = items.map((list: any) => {
      const start = list[this.properties.start];
      const end = list[this.properties.end];

      const bg = (list[this.properties.colorField] || '').toString().trim();
      const bgNormalized = this._normalizeColor(bg);
      const textColor = this._getContrastingTextColor(bgNormalized || '');

      return {
        title: list[this.properties.title],
        start: moment.utc(start, 'YYYY-MM-DD HH:mm:ss').toDate(),
        end: moment.utc(end, 'YYYY-MM-DD HH:mm:ss').toDate(),
        id: list["Id"],
        detail: list[this.properties.detail],
        // Force filled block styling
        display: 'block',
        backgroundColor: bgNormalized,
        borderColor: bgNormalized,
        textColor: textColor,
        classNames: ['gentechEvent'],
        extendedProps: {
          detail: list[this.properties.detail]
        }
      };
    });

    this.context.statusRenderer.clearLoadingIndicator(this.domElement);

    // Create a container for the React Panel if not exists
    let panelContainer = document.getElementById('spfx-calendar-panel');
    if (!panelContainer) {
      panelContainer = document.createElement('div');
      panelContainer.id = 'spfx-calendar-panel';
      this.domElement.appendChild(panelContainer);
    }

    const _renderPanel = (event: IEventDetails | null, isOpen: boolean) => {
      const element = React.createElement(EventDetailsPanel, {
        isOpen,
        onDismiss: () => _renderPanel(null, false),
        event
      });
      ReactDOM.render(element, panelContainer);
    };

    const calendarEl = document.getElementById('spfxcalendar');
    if (calendarEl) {
      const calendar = new Calendar(calendarEl, {
        eventDisplay: 'block',
        displayEventTime: true,
        eventClick: (args: EventClickArg) => {
          const calEvent = args.event;
          // Task B: Open React Panel
          const eventDetails: IEventDetails = {
            title: calEvent.title,
            start: calEvent.start!,
            end: calEvent.end!,
            color: calEvent.backgroundColor, // FullCalendar maps 'color' to 'backgroundColor' usually
            description: calEvent.extendedProps.detail,
            id: calEvent.id,
            url: calEvent.url
          };

          args.jsEvent.preventDefault(); // Prevent default if it's a link
          _renderPanel(eventDetails, true);
        },
        eventDidMount: (info: EventMountArg) => {
          // Task A: Full-row highlight for List View
          if (info.view.type === 'listWeek' || info.view.type === 'listMonth' || info.view.type === 'listYear' || info.view.type === 'listDay') {
            const row = info.el; // The tr element in list view
            const color = info.event.backgroundColor || info.event.borderColor || '#3788d8';

            // Apply left border accent
            row.style.borderLeft = `4px solid ${color}`;

            // Hide the default dot if we use the bar
            const dot = row.querySelector('.fc-list-event-dot') as HTMLElement;
            if (dot) {
              dot.style.display = 'none';
            }
          }
        },
        plugins: [dayGridPlugin, listPlugin],
        initialView: 'dayGridMonth',
        eventSources: [
          {
            events: calItems,
          }
        ],
        headerToolbar: {
          left: 'prev,next today',
          center: 'title',
          right: 'dayGridMonth,listWeek'
        }
      });
      calendar.render();
    }
    //jQuery(".spfxcalendar", this.domElement).fullCalendar(calendarOptions);
  }

  private _getSitesAsync(): void {
    this._getSiteRootWeb().then((response) => {
      this._getSites(response["Url"]).then((response1) => {
        const sites: IPropertyPaneDropdownOption[] = [];
        sites.push({
          key: this.context.pageContext.web.absoluteUrl,
          text: "This Site",
        });
        sites.push({ key: "other", text: "Other Site (Specify Url)" });
        for (const _key in response1.value) {
          sites.push({
            key: response1.value[_key]["Url"],
            text: response1.value[_key]["Title"],
          });
        }
        this._siteOptions = sites;
        this.context.propertyPane.refresh();
        let siteUrl = this.properties.site;
        if (this.properties.other && this.properties.siteOther) {
          siteUrl = this.properties.siteOther
        }
        this._getListTitles(siteUrl).then((response2) => {
          this._dropdownOptions = response2.value.map((list: ISPList) => {
            return {
              key: list.Title,
              text: list.Title,
            };
          });
          this.context.propertyPane.refresh();
          if (this.properties.listTitle) {
            this._log('debug', 'this.properties.site', { site: this.properties.site });
            this._getListColumns(
              this.properties.listTitle,
              this.properties.site
            ).then((response3) => {
              const col: IPropertyPaneDropdownOption[] = [];
              for (const _innerKey in response3.value) {
                col.push({
                  key: response3.value[_innerKey]["InternalName"],
                  text: response3.value[_innerKey]["Title"],
                });
              }
              this._columnOptions = col;
              this.colsDisabled = false;
              this.listDisabled = false;
              this.context.propertyPane.refresh();
              this.context.statusRenderer.clearLoadingIndicator(
                this.domElement
              );
              this.render();
            });
          }
        });
      });
    });
  }

  private _renderListAsync(): void {
    let siteUrl = this.properties.site;
    if (this.properties.other && this.properties.siteOther) {
      siteUrl = this.properties.siteOther;
    }
    this._log('debug', 'siteUrl', { siteUrl });
    this._getListData(this.properties.listTitle!, siteUrl)
      .then((response) => {
        this._log('debug', 'response', { response });
        this._renderList(response.value);
      })
      .catch((err) => {
        this._log('error', 'Error loading list data', err);
        this._renderFatalError(
          "There was an error loading your list. Verify the selected list has Calendar Events or choose a new list.",
          err
        );
      });
  }

  /**
   * Normalizes a color string to a valid CSS color value.
   * check, if the color is a hex value (3 or 6 chars) and add # if missing.
   * check, if the color is a rgb value and add rgb() if missing.
   * @param color The color string to normalize
   */
  private _normalizeColor(color: string): string | undefined {
    if (!color) {
      return undefined;
    }

    color = color.trim();

    // Check for Hex (with or without #)
    // Matches 3 or 6 hex digits, optionally preceded by #
    const hexRegex = /^#?([0-9A-F]{3}|[0-9A-F]{6})$/i;
    if (hexRegex.test(color)) {
      if (!color.startsWith("#")) {
        return "#" + color;
      }
      return color;
    }

    // Check for RGB (e.g., "255,0,0" or "rgb(255,0,0)")
    // loose check for 3 distinct numbers separated by commas
    const rgbRegex = /^(rgb\()?(\d{1,3},\s*\d{1,3},\s*\d{1,3})\)?$/i;
    const match = color.match(rgbRegex);
    if (match) {
      // match[2] contains the numbers part
      return `rgb(${match[2]})`;
    }

    // If it's already a valid named color or other format we don't strictly validate, 
    // we return it as is, or we could return undefined to fallback.
    // For now, let's return it as is if it looks somewhat like a string, 
    // but the requirement was specifically about HEX and RGB.
    // If it fails both strict checks above, we might want to return undefined 
    // to allow FullCalendar to use the default color.
    return undefined;
  }

  private _getContrastingTextColor(hex: string): string {
    if (!hex) return '#000';
    const c = hex.replace('#', '').trim();
    const full = c.length === 3 ? c.split('').map(x => x + x).join('') : c;
    if (full.length !== 6) return '#000';

    const r = parseInt(full.slice(0, 2), 16) / 255;
    const g = parseInt(full.slice(2, 4), 16) / 255;
    const b = parseInt(full.slice(4, 6), 16) / 255;

    // relative luminance
    const lin = (v: number) => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4));
    const L = 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);

    return L > 0.55 ? '#000' : '#fff';
  }
}
