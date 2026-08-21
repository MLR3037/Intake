import { PublicClientApplication } from '@azure/msal-browser';

export function getSharePointConfig() {
  const env = import.meta.env;
  const fallbackRedirect = window.location.origin;
  const siteUrl = env.VITE_SP_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/Clinistrators';
  const hrSiteUrl = env.VITE_SP_HR_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/HR';

  return {
    siteUrl,
    staffSiteUrl: env.VITE_SP_STAFF_SITE_URL || hrSiteUrl,
    staffListName: env.VITE_SP_STAFF_LIST || 'Current Employees',
    // Current clients typically live in EvokeSchedule2.0's own list (often a different list, sometimes a different site).
    currentClientsSiteUrl: env.VITE_SP_CURRENT_CLIENTS_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/EvokeIntake',
    currentClientsListName: env.VITE_SP_CURRENT_CLIENTS_LIST || 'Current Clients',
    dischargedClientsSiteUrl: env.VITE_SP_DISCHARGED_CLIENTS_SITE_URL || env.VITE_SP_CURRENT_CLIENTS_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/EvokeIntake',
    dischargedClientsListName: env.VITE_SP_DISCHARGED_CLIENTS_LIST || 'Discharged Clients',
    // Incoming/prospective clients live in this app's own intake list.
    intakeSiteUrl: env.VITE_SP_INTAKE_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/EvokeIntake',
    intakeListName: env.VITE_SP_INTAKE_LIST || 'Intake Board',
    capacityPlanningSiteUrl: env.VITE_SP_CAPACITY_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/EvokeIntake',
    capacityPlanningListName: env.VITE_SP_CAPACITY_LIST || 'Capacity Planning',
    hrSiteUrl,
    newHireListName: env.VITE_SP_NEW_HIRE_LIST || 'New Hire Onboarding',
    dischargedStaffSiteUrl: env.VITE_SP_DISCHARGED_STAFF_SITE_URL || env.VITE_SP_SITE_URL || 'https://evokebehavioralhealthcom.sharepoint.com/sites/Clinistrators',
    dischargedStaffListName: env.VITE_SP_DISCHARGED_STAFF_LIST || 'Discharged Staff',
    clientId: env.VITE_SP_CLIENT_ID || '',
    tenantId: env.VITE_SP_TENANT_ID || '',
    redirectUri: env.VITE_SP_REDIRECT_URI || fallbackRedirect
  };
}

function normalizeSharePointDate(value) {
  if (!value) return null;
  const raw = String(value);
  const microsoftDate = raw.match(/\/Date\((-?\d+)(?:[+-]\d+)?\)\//);
  if (microsoftDate) return new Date(Number(microsoftDate[1])).toISOString();
  return value;
}

export class SharePointDataService {
  constructor(config) {
    this.config = config;
    this.accessToken = null;
    this.account = null;
    this.msal = null;
    this.initPromise = null;
    this.loginPromise = null;
  }

  isConfigured() {
    return Boolean(this.config.clientId && this.config.tenantId && this.config.siteUrl);
  }

  async initialize() {
    if (!this.isConfigured()) {
      throw new Error('Missing SharePoint config. Set VITE_SP_CLIENT_ID and VITE_SP_TENANT_ID.');
    }

    // Share one in-flight init across concurrent callers instead of racing separate MSAL instances.
    if (!this.initPromise) {
      this.initPromise = this._doInitialize();
    }
    return this.initPromise;
  }

  async _doInitialize() {
    this.msal = new PublicClientApplication({
      auth: {
        clientId: this.config.clientId,
        authority: `https://login.microsoftonline.com/${this.config.tenantId}`,
        redirectUri: this.config.redirectUri
      },
      cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
      }
    });
    await this.msal.initialize();
    await this.msal.handleRedirectPromise();

    const accounts = this.msal.getAllAccounts();
    if (accounts.length > 0) {
      this.account = accounts[0];
    }
  }

  async hasSession() {
    await this.initialize();
    return Boolean(this.account);
  }

  async login() {
    await this.initialize();

    // Redirect flow (matches EvokeSchedule2.0) instead of popup, which can hang if the
    // browser blocks cross-window monitoring of the popup's location.
    await this.msal.loginRedirect({
      scopes: [`${new URL(this.config.siteUrl).origin}/.default`],
      prompt: 'select_account'
    });
  }

  async getAccessToken() {
    if (!this.account) {
      await this.login();
      // loginRedirect navigates away; execution won't reach here until the app reloads.
      return null;
    }

    const request = {
      scopes: [`${new URL(this.config.siteUrl).origin}/.default`],
      account: this.account
    };

    try {
      const silent = await this.msal.acquireTokenSilent(request);
      this.accessToken = silent.accessToken;
      return this.accessToken;
    } catch {
      await this.msal.acquireTokenRedirect(request);
      return null;
    }
  }

  /**
   * Only $select fields that actually exist on the list, so mismatched schemas don't 400.
   * Matching is case-insensitive since SharePoint reports 'ID' while our candidates use 'Id'.
   * 'Id' is always kept since it's a system field that may not appear in the fields list.
   */
  async getAvailableFields(siteUrl, listName, candidateFields) {
    const token = await this.getAccessToken();
    const fieldsUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=InternalName,Title,Hidden&$top=5000`;

    try {
      const response = await fetch(fieldsUrl, {
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=verbose' }
      });
      if (!response.ok) return candidateFields;

      const payload = await response.json();
      const allFields = payload?.d?.results || [];
      const visibleFields = allFields.filter((field) => field.Hidden !== true);

      console.log(
        `\uD83D\uDCC4 '${listName}' available columns:`,
        visibleFields.map((field) => `${field.InternalName} ("${field.Title}")`)
      );

      // Hidden only controls the default SharePoint form, not $select queryability - fields like
      // EndDate commonly come back hidden even when they hold real data, so match against every
      // field returned, not just the visible ones, or a real column silently drops from the query.
      const names = new Set(allFields.map((field) => field.InternalName.toLowerCase()));

      return candidateFields.filter(
        (field) => field === 'Id' || names.has(field.split('/')[0].toLowerCase())
      );
    } catch {
      return candidateFields;
    }
  }

  async getListItems(siteUrl, listName, selectFields, expand) {
    console.log(`📋 Querying SharePoint list '${listName}' at ${siteUrl}`);
    const token = await this.getAccessToken();
    const params = [`$select=${selectFields.join(',')}`, '$top=5000'];
    if (expand) params.push(`$expand=${expand}`);

    const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?${params.join('&')}`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=verbose' }
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Failed to load '${listName}' (${response.status}): ${body}`);
    }

    const payload = await response.json();
    return payload?.d?.results || [];
  }

  async updateListItem(siteUrl, listName, itemId, fields) {
    const token = await this.getAccessToken();
    const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
    const listResponse = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=verbose' } }
    );
    if (!listResponse.ok) {
      const body = await listResponse.text();
      throw new Error(`Failed to inspect '${listName}' (${listResponse.status}): ${body}`);
    }
    const listPayload = await listResponse.json();
    const entityType = listPayload?.d?.ListItemEntityTypeFullName;
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        __metadata: { type: entityType },
        ...fields
      })
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Failed to update '${listName}' (${response.status}): ${body}`);
    }
  }

  async createListItem(siteUrl, listName, fields) {
    const token = await this.getAccessToken();
    const listResponse = await fetch(
      `${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json;odata=verbose' } }
    );
    if (!listResponse.ok) {
      const body = await listResponse.text();
      throw new Error(`Failed to inspect '${listName}' (${listResponse.status}): ${body}`);
    }
    const listPayload = await listResponse.json();
    const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify({
        __metadata: { type: listPayload?.d?.ListItemEntityTypeFullName },
        ...fields
      })
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Failed to create '${listName}' (${response.status}): ${body}`);
    }

    const payload = await response.json();
    return payload?.d;
  }

  async loadStaff(listName = this.config.staffListName, siteUrl = this.config.staffSiteUrl, forceInactive = false) {
    const candidateFields = [
      'Id',
      'StaffPerson/Id',
      'StaffPerson/Title',
      'StaffPerson/EMail',
      'Title',
      'Name',
      'Role',
      'Position',
      'IsActive',
      'StartDate',
      'Start Date',
      'Start_x0020_Date',
      'TermDate',
      'Term Date',
      'Term_x0020_Date',
      'NewPosition',
      'New Position',
      'New_x0020_Position',
      'NewPositionStartDate',
      'New Position Start Date',
      'NewPositionStart_x0020_Date'
    ];
    const select = await this.getAvailableFields(siteUrl, listName, candidateFields);
    const expand = select.some((field) => field.startsWith('StaffPerson/')) ? 'StaffPerson' : undefined;
    const items = await this.getListItems(siteUrl, listName, select, expand);

    return items.map((item) => {
      const person = item.StaffPerson || {};
      const readItemField = (names) => {
        const matchingName = Object.keys(item).find((key) => names.some((name) => key.toLowerCase() === name.toLowerCase()));
        return matchingName ? item[matchingName] : null;
      };
      const readSharePointValue = (value) => {
        if (value && typeof value === 'object') {
          return value.Value || value.Label || value.Title || value.lookupValue || '';
        }
        return value || '';
      };
      const role = normalizeStaffPosition(readSharePointValue(readItemField(['Position', 'Role']) || 'RBT'));
      const newPosition = String(readSharePointValue(readItemField(['NewPosition', 'New Position', 'New_x0020_Position']))).trim();
      const newPositionStartDate = normalizeSharePointDate(
        readSharePointValue(readItemField(['NewPositionStartDate', 'New Position Start Date', 'NewPositionStart_x0020_Date'])) || null
      );
      const transitionEndDate = isDirectCareProjectionRole(role) &&
        newPosition &&
        !isBehaviorSpecialistPosition(newPosition) &&
        newPositionStartDate
        ? newPositionStartDate
        : null;
      return {
        id: person.Id || item.Id,
        sourceId: item.Id,
        name: person.Title || readSharePointValue(readItemField(['Title', 'Name'])) || 'Unknown Staff',
        email: person.EMail || '',
        role,
        startDate: normalizeSharePointDate(readSharePointValue(readItemField(['StartDate', 'Start Date', 'Start_x0020_Date'])) || null),
        endDate: transitionEndDate || normalizeSharePointDate(readSharePointValue(readItemField(['TermDate', 'Term Date', 'Term_x0020_Date'])) || null),
        newPosition,
        newPositionStartDate,
        isActive: forceInactive ? false : item.IsActive !== false
      };
    });
  }

  async loadDischargedStaff() {
    try {
      return await this.loadStaff(this.config.dischargedStaffListName, this.config.dischargedStaffSiteUrl, true);
    } catch (error) {
      console.warn(`Unable to load optional discharged staff list '${this.config.dischargedStaffListName}':`, error);
      return [];
    }
  }

  async loadCurrentClients(listName = this.config.currentClientsListName, siteUrl = this.config.currentClientsSiteUrl, forceInactive = false) {
    const candidateFields = [
      'Id',
      'Title',
      'UpdatedProgram',
      'Service',
      'RatioAM',
      'IsActive',
      'IntakeStatus',
      'Intake_x0020_Status',
      'StartDate',
      'EndDate'
    ];
    const select = await this.getAvailableFields(
      siteUrl,
      listName,
      candidateFields
    );
    const items = await this.getListItems(
      siteUrl,
      listName,
      select
    );

    return items.map((item) => {
      const intakeStatus = item.IntakeStatus || item.Intake_x0020_Status || '';
      const isDiscontinued = String(intakeStatus).trim().toLowerCase() === 'discontinued';
      const dischargeDate = normalizeSharePointDate(item.EndDate || null);
      const hasDischarged = dischargeDate && new Date(dischargeDate) < new Date();

      return {
        id: item.Id,
        sourceId: item.Id,
        Title: item.Title || `Client ${item.Id}`,
        Program: item.UpdatedProgram || '',
        Services: item.Service || '',
        'Start Date': normalizeSharePointDate(item.StartDate),
        'Discharge Date': dischargeDate,
        'Intake Status': intakeStatus,
        'Staffing Ratio': item.RatioAM || '1:1',
        Status: forceInactive || item.IsActive === false || (isDiscontinued && hasDischarged) ? 'Inactive' : 'Active'
      };
    });
  }

  async loadDischargedClients() {
    try {
      return await this.loadCurrentClients(
        this.config.dischargedClientsListName,
        this.config.dischargedClientsSiteUrl,
        true
      );
    } catch (error) {
      console.warn(`Unable to load optional discharged client list '${this.config.dischargedClientsListName}':`, error);
      return [];
    }
  }

  async loadIntakeClients() {
    const candidateFields = [
      'Id',
      'Title',
      'board_x0020_choice',
      'Services',
      'InquiryDate',
      'ReferralSource',
      'SchoolDistrict',
      'StartDate',
      'TentativeStartDate',
      'StaffingRatio'
    ];
    const select = await this.getAvailableFields(this.config.intakeSiteUrl, this.config.intakeListName, candidateFields);
    const items = await this.getListItems(this.config.intakeSiteUrl, this.config.intakeListName, select);

    return items.map((item) => ({
      id: item.Id,
      Title: item.Title || `Prospect ${item.Id}`,
      'Intake Status': item.board_x0020_choice || 'Initial Inquiry',
      Services: item.Services || 'ABA',
      'Inquiry Date': normalizeSharePointDate(item.InquiryDate),
      'Referral Source': item.ReferralSource || '',
      'School District': item.SchoolDistrict || '',
      'Start Date': normalizeSharePointDate(item.StartDate),
      'Tentative Start Date': normalizeSharePointDate(item.TentativeStartDate),
      'Staffing Ratio': item.StaffingRatio || '1:1'
    }));
  }

  async loadProspectiveStaff() {
    const items = await this.getListItems(
      this.config.hrSiteUrl,
      this.config.newHireListName,
      ['Id', 'Title', 'Position', 'TermDate', 'HireType']
    );

    const readSharePointValue = (value) => {
      if (value && typeof value === 'object') {
        return value.Value || value.Label || value.Title || value.lookupValue || '';
      }
      return value || '';
    };

    return items
      .map((item) => {
        const position = normalizeStaffPosition(readSharePointValue(item.Position));
        const name = readSharePointValue(item.Title) || `New Hire ${item.Id}`;
        const startDate = normalizeSharePointDate(readSharePointValue(item.TermDate));
        const hireType = String(readSharePointValue(item.HireType)).trim();

        return {
          id: item.Id,
          sourceId: item.Id,
          name,
          position: position.toString().trim(),
          startDate,
          hireType,
          status: 'Upcoming'
        };
      })
      .filter((person) => person.startDate && person.hireType.toLowerCase() === 'external');
  }

  async loadCapacityPlanning() {
    const fields = [
      'Id',
      'Title',
      'SourceList',
      'SourceID',
      'PersonType',
      'MonAM',
      'MonPM',
      'TueAM',
      'TuePM',
      'WedAM',
      'WedPM',
      'ThuAM',
      'ThuPM',
      'FriAM',
      'FriPM'
    ];
    const select = await this.getAvailableFields(
      this.config.capacityPlanningSiteUrl,
      this.config.capacityPlanningListName,
      fields
    );
    const items = await this.getListItems(
      this.config.capacityPlanningSiteUrl,
      this.config.capacityPlanningListName,
      select
    );

    const readPlanningValue = (value) => {
      if (value && typeof value === 'object') {
        return value.Value || value.Label || value.Title || value.lookupValue || '';
      }
      return value || '';
    };

    return items.map((item) => ({
      id: item.Id,
      title: readPlanningValue(item.Title),
      sourceList: readPlanningValue(item.SourceList),
      sourceId: String(readPlanningValue(item.SourceID)),
      personType: readPlanningValue(item.PersonType),
      availability: Object.fromEntries(
        ['MonAM', 'MonPM', 'TueAM', 'TuePM', 'WedAM', 'WedPM', 'ThuAM', 'ThuPM', 'FriAM', 'FriPM'].map((field) => [
          field,
          item[field] === true
        ])
      )
    }));
  }

  async loadSnapshot() {
    // Authenticate once before firing the four list loads in parallel, so they share one session instead of racing.
    await this.getAccessToken();

    const [staff, dischargedStaff, currentClients, dischargedClients, intakeClients, prospectiveStaff, capacityPlanning] = await Promise.all([
      this.loadStaff(),
      this.loadDischargedStaff(),
      this.loadCurrentClients(),
      this.loadDischargedClients(),
      this.loadIntakeClients(),
      this.loadProspectiveStaff(),
      this.loadCapacityPlanning()
    ]);

    return {
      staffing: summarizeStaffing([...staff, ...dischargedStaff]),
      currentClients: [...currentClients, ...dischargedClients],
      intakeClients,
      prospectiveStaff,
      capacityPlanning,
      totalClients: currentClients.length + intakeClients.length,
      lastUpdatedAt: new Date().toISOString()
    };
  }
}

export function summarizeStaffing(staff) {
  const activeStaff = staff.filter((person) => person.isActive);
  const availableDirectStaff = activeStaff.filter((person) => isAvailableDirectStaffRole(person.role));

  const byRole = activeStaff.reduce((counts, person) => {
    counts[person.role] = (counts[person.role] || 0) + 1;
    return counts;
  }, {});

  return {
    staff,
    activeCount: availableDirectStaff.length,
    totalActiveCount: activeStaff.length,
    byRole,
    lastUpdatedAt: new Date().toISOString()
  };
}

export function isAvailableDirectStaffRole(value) {
  const normalized = String(value || '')
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  return (
    normalized === 'rbt' ||
    normalized === 'bt' ||
    normalized === 'bt1' ||
    normalized === 'bt 1' ||
    normalized === 'behavior technician 1' ||
    normalized === 'behavior technician i' ||
    normalized === 'bt2' ||
    normalized === 'bt 2' ||
    normalized === 'behavior technician 2' ||
    normalized === 'behavior technician ii'
  );
}

function isDirectCareProjectionRole(value) {
  const normalized = String(value || '').toLowerCase().trim();
  return isAvailableDirectStaffRole(value) || normalized === 'specialist' || normalized === 'specialists' || normalized.includes('behavior specialist');
}

function isBehaviorSpecialistPosition(value) {
  return String(value || '').toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim().includes('behavior specialist');
}

export function normalizeStaffPosition(value) {
  const normalized = String(value || '')
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (
    normalized === 'rbt' ||
    normalized === 'bt' ||
    normalized === 'bt1' ||
    normalized === 'bt 1' ||
    normalized === 'behavior technician 1' ||
    normalized === 'behavior technician i' ||
    normalized === 'bt2' ||
    normalized === 'bt 2' ||
    normalized === 'behavior technician 2' ||
    normalized === 'behavior technician ii'
  ) {
    return 'BT';
  }

  if (normalized.includes('behavior specialist')) return 'Behavior Specialist';
  return String(value || '').trim();
}

export function parseStaffUnits(ratio) {
  if (!ratio || typeof ratio !== 'string') return 1;

  const [staffCountRaw, clientCountRaw] = ratio.split(':');
  const staffCount = Number.parseFloat(staffCountRaw);
  const clientCount = Number.parseFloat(clientCountRaw);

  if (!Number.isFinite(staffCount) || !Number.isFinite(clientCount) || clientCount <= 0) {
    return 1;
  }

  return staffCount / clientCount;
}

export const REQUIRED_STAFF_PER_CLIENT = 1.125;
export const TWO_TO_ONE_STAFF_PER_CLIENT = 2;
export const RESERVED_CLIENT_SPOTS = 1;
export const STAFF_TRAINING_PERIOD_DAYS = 21;

export function getRequiredStaffUnits(client) {
  return parseStaffUnits(client['Staffing Ratio']) === TWO_TO_ONE_STAFF_PER_CLIENT
    ? TWO_TO_ONE_STAFF_PER_CLIENT
    : REQUIRED_STAFF_PER_CLIENT;
}

export function calculateRequiredStaff(clients) {
  const total = clients.reduce(
    (sum, client) => sum + getRequiredStaffUnits(client),
    RESERVED_CLIENT_SPOTS
  );
  return Number(total.toFixed(2));
}

function asCalendarDate(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

export function generateCapacityProjection(
  currentClients,
  intakeData,
  activeStaffCount,
  prospectiveStaff = [],
  weekCount = 8,
  staffRecords = [],
  historyWeeks = 4
) {
  const weeks = [];
  const today = new Date();
  const todayDate = asCalendarDate(today);

  for (let index = -historyWeeks; index < weekCount; index += 1) {
    const weekStart = new Date(today);
    weekStart.setDate(weekStart.getDate() + index * 7);
    weekStart.setHours(0, 0, 0, 0);
    const weekLabel = weekStart.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
    const weekStartDate = asCalendarDate(weekStart);
    const isHistorical = weekStartDate < todayDate;

    const activeCurrentClients = currentClients.filter((client) => {
      const startDate = client['Start Date'] ? new Date(client['Start Date']) : null;
      const dischargeDate = client['Discharge Date'] ? new Date(client['Discharge Date']) : null;

      if (client.Status === 'Inactive' && (!dischargeDate || weekStartDate > asCalendarDate(dischargeDate))) return false;

      if (startDate && asCalendarDate(startDate) > weekStartDate) return false;
      if (dischargeDate && weekStartDate > asCalendarDate(dischargeDate)) return false;
      return true;
    });

    const expectedEnrollments = intakeData.filter((item) => {
      const startDateValue = item['Start Date'] || item['Tentative Start Date'];
      if (!startDateValue) return false;
      const startDate = new Date(startDateValue);
      return startDate <= weekStart;
    });

    const requiredStaff = calculateRequiredStaff([...activeCurrentClients, ...expectedEnrollments]);
    const staffInTraining = prospectiveStaff.filter(
      (person) =>
        isAvailableDirectStaffRole(person.position) &&
        person.startDate &&
        new Date(person.startDate) <= weekStart &&
        new Date(new Date(person.startDate).setDate(new Date(person.startDate).getDate() + STAFF_TRAINING_PERIOD_DAYS)) > weekStart
    ).length;
    const onboardedByWeek = prospectiveStaff.filter(
      (person) =>
        isAvailableDirectStaffRole(person.position) &&
        person.startDate &&
        new Date(new Date(person.startDate).setDate(new Date(person.startDate).getDate() + STAFF_TRAINING_PERIOD_DAYS)) <= weekStart
    ).length;
    const availableDirectStaffByWeek = staffRecords.filter((person) => {
      if (!person || (person.isActive === false && !person.endDate)) return false;
      if (!isAvailableDirectStaffRole(person.role)) return false;

      const startDate = person.startDate ? asCalendarDate(person.startDate) : null;
      const endDate = person.endDate ? asCalendarDate(person.endDate) : null;

      if (startDate && startDate > weekStart) return false;
      if (endDate && weekStart > endDate) return false;
      return true;
    }).length;
    // Base available staff reflects today's direct-care roster, minus any staff already scheduled
    // to end before that future week. New hires are then layered on top as they reach their start date.
    const availableStaff = staffRecords.length > 0 ? availableDirectStaffByWeek : activeStaffCount || 0;
    const projectedStudents = activeCurrentClients.length + expectedEnrollments.length;
    const projectedStaff = availableStaff + onboardedByWeek;
    const staffingGap = Number((projectedStaff - requiredStaff).toFixed(2));

    weeks.push({
      week: weekLabel,
      currentClients: activeCurrentClients.length,
      projectedStudents,
      requiredStaff,
      availableStaff,
      staffInTraining,
      projectedStaff,
      staffingGap,
      isHistorical
    });
  }

  return weeks;
}

export function applyWhatIfAdjustments(weeks, { staffDelta = 0, clientDelta = 0 } = {}) {
  if (!staffDelta && !clientDelta) return weeks;

  return weeks.map((week) => {
    if (week.isHistorical) return week;

    const projectedStudents = Math.max(0, week.projectedStudents + clientDelta);
    const requiredStaff = Number(
      (projectedStudents * REQUIRED_STAFF_PER_CLIENT + RESERVED_CLIENT_SPOTS).toFixed(2)
    );
    // staffDelta is a hypothetical - it never rewrites today's real availableStaff, only the
    // forward-looking projectedStaff line (and the gap that depends on it).
    const projectedStaff = Math.max(0, week.projectedStaff + staffDelta);
    const staffingGap = Number((projectedStaff - requiredStaff).toFixed(2));

    return {
      ...week,
      projectedStudents,
      requiredStaff,
      projectedStaff,
      staffingGap
    };
  });
}

function formatDateForExport(value) {
  if (!value) return '';
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return value;

  return parsed.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric',
    year: 'numeric'
  });
}

export function exportIntakeDataToCSV(data) {
  if (!data.length) {
    window.alert('No data to export');
    return;
  }

  const headers = [
    'Client Name',
    'Service',
    'Inquiry Date',
    'Intake Status',
    'Tentative Start',
    'Staffing Ratio',
    'Days to Start',
    'Referral Source',
    'School District',
    'Notes'
  ];

  const rows = data.map((item) => [
    item.Title,
    item.Services,
    formatDateForExport(item['Inquiry Date']),
    item['Intake Status'],
    formatDateForExport(item['Tentative Start Date']),
    item['Staffing Ratio'],
    item.DaysToStart,
    item['Referral Source'],
    item['School District'],
    item.Notes || ''
  ]);

  const csv = [headers.join(','), ...rows.map((row) => row.map((cell) => `"${cell}"`).join(','))].join('\n');

  const blob = new Blob([csv], { type: 'text/csv' });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `intake-report-${new Date().toISOString().split('T')[0]}.csv`;
  anchor.click();
  window.URL.revokeObjectURL(url);
}
