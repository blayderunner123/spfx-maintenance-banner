import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import * as strings from 'MaintenanceBannerApplicationCustomizerStrings';

const LOG_SOURCE = 'MaintenanceBannerApplicationCustomizer';
const CONFIGURATION_LIST = 'Site Maintenance Banners';
const DISMISS_KEY_PREFIX = 'spfx-maintenance-banner-dismissed-';

export interface IMaintenanceBannerApplicationCustomizerProperties {
  /** Absolute URL of the web containing the configuration list. Defaults to the current web. */
  configurationSiteUrl?: string;
}

interface IAudienceValue {
  Id: number;
  Title: string;
}

interface IBannerItem {
  Id: number;
  Title?: string;
  Message?: string;
  Enabled?: boolean;
  BannerType?: string;
  StartDate?: string;
  EndDate?: string;
  TargetSites?: string;
  ExcludedSites?: string;
  TargetUsers?: string;
  TargetAudience?: IAudienceValue[];
  Priority?: number;
  LinkText?: string;
  LinkUrl?: string | { Url?: string };
  Dismissible?: boolean;
  ShowCountdown?: boolean;
  CountdownLabel?: string;
  OpenLinkInNewTab?: boolean;
  AutoRefreshSeconds?: number;
}

interface IGraphGroup {
  id?: string;
  displayName?: string;
  mail?: string;
}

interface IGraphGroupPage {
  value?: IGraphGroup[];
  '@odata.nextLink'?: string;
}

export default class MaintenanceBannerApplicationCustomizer
  extends BaseApplicationCustomizer<IMaintenanceBannerApplicationCustomizerProperties> {
  private _sp!: SPFI;
  private _topPlaceholder?: PlaceholderContent;
  private _refreshTimer?: number;
  private _currentUserSharePointGroupTitles?: Set<string>;
  private _currentUserSharePointGroupIds?: Set<number>;
  private _currentUserGraphAudienceValues?: Set<string>;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    const configurationSiteUrl = this.properties.configurationSiteUrl ||
      this.context.pageContext.web.absoluteUrl;
    this._sp = spfi(configurationSiteUrl).using(SPFx(this.context));
    this._renderBanner().catch((error: unknown) => console.error('Banner initialization error:', error));
    return Promise.resolve();
  }

  private async _renderBanner(): Promise<void> {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
      if (!this._topPlaceholder) {
        console.error('Top placeholder not found.');
        return;
      }
    }

    try {
      const banners: IBannerItem[] = await this._sp.web.lists
        .getByTitle(CONFIGURATION_LIST).items
        .select(
          'Id', 'Title', 'Message', 'Enabled', 'BannerType', 'StartDate', 'EndDate',
          'TargetSites', 'ExcludedSites', 'TargetUsers', 'TargetAudience/Id',
          'TargetAudience/Title', 'Priority', 'LinkText', 'LinkUrl', 'Dismissible',
          'ShowCountdown', 'CountdownLabel', 'OpenLinkInNewTab', 'AutoRefreshSeconds'
        )
        .expand('TargetAudience')
        .filter('Enabled eq 1')
        .orderBy('Priority', true)
        .top(5)();

      let html = '';
      const now = new Date();
      const absoluteUrl = this.context.pageContext.web.absoluteUrl.toLowerCase();
      const serverRelativeUrl = this.context.pageContext.web.serverRelativeUrl.toLowerCase();
      const pathname = window.location.pathname.toLowerCase();
      const userValues = [
        this.context.pageContext.user.email,
        this.context.pageContext.user.loginName,
        this.context.pageContext.user.displayName
      ].filter((value): value is string => !!value).map(value => value.toLowerCase());

      for (const banner of banners) {
        if ((banner.StartDate && new Date(banner.StartDate) > now) ||
            (banner.EndDate && new Date(banner.EndDate) < now)) continue;
        if (!this._targetSitesMatch(banner.TargetSites, absoluteUrl, serverRelativeUrl, pathname)) continue;
        if (this._excludedSitesMatch(banner.ExcludedSites, absoluteUrl, serverRelativeUrl, pathname)) continue;
        if (!await this._targetAudienceAllowed(banner, userValues)) continue;

        const dismissKey = `${DISMISS_KEY_PREFIX}${banner.Id}`;
        if (banner.Dismissible && localStorage.getItem(dismissKey) === 'true') continue;

        const color = this._getBannerColor(banner.BannerType || 'Info');
        const linkUrl = this._getLinkUrl(banner.LinkUrl);
        const linkText = banner.LinkText || 'View Details';
        const linkTarget = banner.OpenLinkInNewTab ? ' target="_blank" rel="noopener noreferrer"' : '';
        const countdown = this._getCountdownText(banner.EndDate || '', !!banner.ShowCountdown, banner.CountdownLabel || '');
        const link = linkUrl ? `<a href="${linkUrl}"${linkTarget} style="margin-left: 12px; color: white; text-decoration: underline; font-weight: 700; white-space: nowrap;">${linkText}</a>` : '';
        const dismiss = banner.Dismissible ? `<button type="button" data-banner-dismiss="${banner.Id}" aria-label="Dismiss banner" style="position: absolute; right: 12px; top: 50%; transform: translateY(-50%); background: transparent; border: none; color: white; cursor: pointer; font-size: 16px; font-weight: 600; line-height: 1; opacity: 0.85; padding: 4px 8px; transition: opacity 0.2s ease-in-out;" onmouseover="this.style.opacity='1'" onmouseout="this.style.opacity='0.85'">&#10005;</button>` : '';

        html += `<div data-banner-id="${banner.Id}" style="width: 100%; background-color: ${color}; color: white; padding: 10px 42px 10px 42px; text-align: center; font-weight: 600; border-bottom: 1px solid rgba(0, 0, 0, 0.2); position: relative; box-sizing: border-box;">${banner.Message || ''}${countdown}${link}${dismiss}</div>`;
      }

      this._topPlaceholder.domElement.innerHTML = html;
      this._bindDismissButtons();
      this._configureAutoRefresh(banners);
    } catch (error) {
      console.error('Banner load error:', error);
    }
  }

  private async _targetAudienceAllowed(banner: IBannerItem, userValues: string[]): Promise<boolean> {
    const directUsers = this._splitValues(banner.TargetUsers);
    const audiences = banner.TargetAudience || [];
    if (!directUsers.length && !audiences.length) return true;
    if (directUsers.length && this._targetUsersMatch(banner.TargetUsers, userValues)) return true;
    return audiences.length ? this._targetAudienceMatch(audiences, userValues) : false;
  }

  private async _targetAudienceMatch(audiences: IAudienceValue[], userValues: string[]): Promise<boolean> {
    await this._ensureCurrentUserSharePointGroups();
    await this._ensureCurrentUserGraphAudienceValues();
    for (const audience of audiences) {
      const title = (audience.Title || '').toLowerCase();
      const id = Number(audience.Id);
      if (title && userValues.some(value => value.indexOf(title) !== -1)) return true;
      if (title && this._currentUserSharePointGroupTitles?.has(title)) return true;
      if (id && this._currentUserSharePointGroupIds?.has(id)) return true;
      if (title && this._currentUserGraphAudienceValues?.has(title)) return true;
    }
    return false;
  }

  private async _ensureCurrentUserSharePointGroups(): Promise<void> {
    if (this._currentUserSharePointGroupTitles && this._currentUserSharePointGroupIds) return;
    this._currentUserSharePointGroupTitles = new Set<string>();
    this._currentUserSharePointGroupIds = new Set<number>();
    try {
      const groups: IAudienceValue[] = await this._sp.web.currentUser.groups.select('Id', 'Title')();
      groups.forEach(group => {
        this._currentUserSharePointGroupIds?.add(group.Id);
        if (group.Title) this._currentUserSharePointGroupTitles?.add(group.Title.toLowerCase());
      });
    } catch (error) {
      console.error('Unable to load current SharePoint group membership:', error);
    }
  }

  private async _ensureCurrentUserGraphAudienceValues(): Promise<void> {
    if (this._currentUserGraphAudienceValues) return;
    this._currentUserGraphAudienceValues = new Set<string>();
    try {
      const client = await this.context.msGraphClientFactory.getClient('3');
      let page: IGraphGroupPage | undefined = await client.api('/me/transitiveMemberOf/microsoft.graph.group')
        .select('id,displayName,mail').top(999).get();
      while (page) {
        page.value?.forEach(group => {
          if (group.id) this._currentUserGraphAudienceValues?.add(group.id.toLowerCase());
          if (group.displayName) this._currentUserGraphAudienceValues?.add(group.displayName.toLowerCase());
          if (group.mail) this._currentUserGraphAudienceValues?.add(group.mail.toLowerCase());
        });
        page = page['@odata.nextLink'] ? await client.api(page['@odata.nextLink']).get() : undefined;
      }
    } catch (error) {
      console.error('Unable to load current Graph group membership:', error);
    }
  }

  private _getBannerColor(type: string): string {
    switch (type) {
      case 'Critical': return '#a4262c';
      case 'Warning': return '#f9a825';
      case 'Success': return '#107c10';
      default: return '#20558a';
    }
  }

  private _getLinkUrl(value?: string | { Url?: string }): string {
    return value ? (typeof value === 'string' ? value : value.Url || '') : '';
  }

  private _getCountdownText(endDate: string, enabled: boolean, label: string): string {
    if (!enabled || !endDate) return '';
    const remaining = new Date(endDate).getTime() - Date.now();
    if (remaining <= 0) return ' | Expired';
    const days = Math.floor(remaining / 86400000);
    const hours = Math.floor((remaining / 3600000) % 24);
    const minutes = Math.floor((remaining / 60000) % 60);
    return ` | ${label || 'Ends in'} ${days}d ${hours}h ${minutes}m`;
  }

  private _splitValues(value?: string): string[] {
    return value && value.trim() !== ''
      ? value.toLowerCase().split(/\r?\n|,|;/).map(item => item.trim()).filter(item => item.length > 0)
      : [];
  }

  private _targetSitesMatch(value: string | undefined, absoluteUrl: string, serverRelativeUrl: string, pathname: string): boolean {
    const targets = this._splitValues(value);
    return !targets.length || targets.some(target => absoluteUrl.indexOf(target) !== -1 || serverRelativeUrl.indexOf(target) !== -1 || pathname.indexOf(target) !== -1);
  }

  private _excludedSitesMatch(value: string | undefined, absoluteUrl: string, serverRelativeUrl: string, pathname: string): boolean {
    const exclusions = this._splitValues(value);
    return !!exclusions.length && exclusions.some(target => absoluteUrl.indexOf(target) !== -1 || serverRelativeUrl.indexOf(target) !== -1 || pathname.indexOf(target) !== -1);
  }

  private _targetUsersMatch(value: string | undefined, userValues: string[]): boolean {
    const targets = this._splitValues(value);
    return !!targets.length && targets.some(target => userValues.some(user => user.indexOf(target) !== -1));
  }

  private _configureAutoRefresh(banners: IBannerItem[]): void {
    const intervals = banners.map(item => Number(item.AutoRefreshSeconds)).filter(value => !!value && value >= 30);
    if (intervals.length) {
      const seconds = Math.min(...intervals);
      if (this._refreshTimer) window.clearInterval(this._refreshTimer);
      this._refreshTimer = window.setInterval(() => {
        this._renderBanner().catch(error => console.error('Banner refresh error:', error));
      }, seconds * 1000);
    }
  }

  private _bindDismissButtons(): void {
    this._topPlaceholder?.domElement.querySelectorAll('[data-banner-dismiss]').forEach(button => {
      button.addEventListener('click', () => {
        const id = button.getAttribute('data-banner-dismiss');
        if (!id) return;
        localStorage.setItem(`${DISMISS_KEY_PREFIX}${id}`, 'true');
        const banner = this._topPlaceholder?.domElement.querySelector(`[data-banner-id="${id}"]`);
        if (banner?.parentElement) banner.parentElement.removeChild(banner);
      });
    });
  }

  private _onDispose = (): void => {
    if (this._refreshTimer) window.clearInterval(this._refreshTimer);
    console.log('Disposed custom banner.');
  };
}
/*
 * SharePoint Maintenance Banner
 * Copyright © 2026 Jonathan R. Adcox
 * Original author: Jonathan R. Adcox (Blayderunner123)
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */
