import { dateAdd, hOP } from "@pnp/core";
import {
    _OLD_SharePointQueryableInstance,
    OLD_SharePointQueryableCollection,
    OLD_ISharePointQueryableCollection,
    _OLD_SharePointQueryableCollection,
    OLD_spInvokableFactory,
} from "../sharepointqueryable.js";
import { defaultPath } from "../decorators.js";
import { spODataEntity } from "../odata.js";
import { OLD_spPost } from "../operations.js";
import { tag } from "../telemetry.js";

@defaultPath("regionalsettings")
export class _RegionalSettings extends _OLD_SharePointQueryableInstance<IRegionalSettingsInfo> {

    /**
     * Gets the collection of languages used in a server farm.
     * ** Please use getInstalledLanguages instead of this method **
     */
    public get installedLanguages(): OLD_ISharePointQueryableCollection<{ Items: IInstalledLanguageInfo[] }> {
        console.warn("Deprecated: RegionalSettings.installedLanguages is deprecated, please use RegionalSettings.getInstalledLanguages");
        return <any>tag.configure(OLD_SharePointQueryableCollection(this, "installedlanguages"), "rs.installedLanguages");
    }

    /**
     * Gets time zone
     */
    public get timeZone(): ITimeZone {
        return tag.configure(TimeZone(this), "rs.tz");
    }

    /**
     * Gets time zones
     */
    public get timeZones(): ITimeZones {
        return tag.configure(TimeZones(this), "rs.tzs");
    }

    /**
     * Gets the collection of languages used in a server farm.
     */
    public async getInstalledLanguages(): Promise<IInstalledLanguageInfo[]> {
        const results: { Items: IInstalledLanguageInfo[] } = await tag.configure(OLD_SharePointQueryableCollection(this, "installedlanguages"), "rs.getInstalledLanguages")();
        return results.Items;
    }
}
export interface IRegionalSettings extends _RegionalSettings {}
export const RegionalSettings = OLD_spInvokableFactory<IRegionalSettings>(_RegionalSettings);

@defaultPath("timezone")
export class _TimeZone extends _OLD_SharePointQueryableInstance<ITimeZoneInfo> {

    /**
     * Gets an Local Time by UTC Time
     *
     * @param utcTime UTC Time as Date or ISO String
     */
    @tag("tz.utcToLocalTime")
    public async utcToLocalTime(utcTime: string | Date): Promise<string> {

        let dateIsoString: string;

        if (typeof utcTime === "string") {
            dateIsoString = utcTime;
        } else {
            dateIsoString = utcTime.toISOString();
        }

        const res = await OLD_spPost(this.clone(TimeZone, `utctolocaltime('${dateIsoString}')`));
        return hOP(res, "UTCToLocalTime") ? res.UTCToLocalTime : res;
    }

    /**
     * Gets an UTC Time by Local Time
     *
     * @param localTime Local Time as Date or ISO String
     */
    @tag("tz.localTimeToUTC")
    public async localTimeToUTC(localTime: string | Date): Promise<string> {

        let dateIsoString: string;

        if (typeof localTime === "string") {
            dateIsoString = localTime;
        } else {
            dateIsoString = dateAdd(localTime, "minute", localTime.getTimezoneOffset() * -1).toISOString();
        }

        const res = await OLD_spPost(this.clone(TimeZone, `localtimetoutc('${dateIsoString}')`));

        return hOP(res, "LocalTimeToUTC") ? res.LocalTimeToUTC : res;
    }
}
export interface ITimeZone extends _TimeZone {}
export const TimeZone = OLD_spInvokableFactory<ITimeZone>(_TimeZone);

@defaultPath("timezones")
export class _TimeZones extends _OLD_SharePointQueryableCollection<ITimeZoneInfo[]> {

    /**
     * Gets an TimeZone by id (see: https://msdn.microsoft.com/en-us/library/office/jj247008.aspx)
     *
     * @param id The integer id of the timezone to retrieve
     */
    @tag("tzs.getById")
    public getById(id: number): Promise<ITimeZone & ITimeZoneInfo> {
        // do the post and merge the result into a TimeZone instance so the data and methods are available
        return OLD_spPost(this.clone(TimeZones, `GetById(${id})`).usingParser(spODataEntity(TimeZone)));
    }
}
export interface ITimeZones extends _TimeZones {}
export const TimeZones = OLD_spInvokableFactory<ITimeZones>(_TimeZones);

/**
 * This is the data for Regional Settings
 */
export interface IRegionalSettingsInfo {
    AdjustHijriDays: number;
    AlternateCalendarType: number;
    AM: string;
    CalendarType: number;
    Collation: number;
    CollationLCID: number;
    DateFormat: number;
    DateSeparator: string;
    DecimalSeparator: string;
    DigitGrouping: string;
    FirstDayOfWeek: number;
    FirstWeekOfYear: number;
    IsEastAsia: boolean;
    IsRightToLeft: boolean;
    IsUIRightToLeft: boolean;
    ListSeparator: string;
    LocaleId: number;
    NegativeSign: string;
    NegNumberMode: number;
    PM: string;
    PositiveSign: string;
    ShowWeeks: boolean;
    ThousandSeparator: string;
    Time24: boolean;
    TimeMarkerPosition: number;
    TimeSeparator: string;
    WorkDayEndHour: number;
    WorkDays: number;
    WorkDayStartHour: number;
}

export interface IInstalledLanguageInfo {
    DisplayName: string;
    LanguageTag: string;
    Lcid: number;
}

export interface ITimeZoneInfo {
    Description: string;
    Id: number;
    Information: {
        Bias: number;
        DaylightBias: number;
        StandardBias: number;
    };
}

export interface IUserResources {
    /**
     * Gets the resource string for the title
     */
    titleResource(cultureName: string): Promise<string>;
    /**
     * Gets the resource string for the title description
     */
    descriptionResource(cultureName: string): Promise<string>;
}
