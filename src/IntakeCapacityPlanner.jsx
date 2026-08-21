import React, { useEffect, useRef, useState } from 'react';
import {
  Area,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ComposedChart,
  Legend,
  Line,
  LabelList,
  Pie,
  PieChart,
  ReferenceLine,
  ReferenceArea,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis
} from 'recharts';
import { AlertCircle, ChevronDown, Download, Filter } from 'lucide-react';
import {
  applyWhatIfAdjustments,
  exportIntakeDataToCSV,
  generateCapacityProjection,
  getSharePointConfig,
  isAvailableDirectStaffRole,
  getRequiredStaffUnits,
  REQUIRED_STAFF_PER_CLIENT,
  RESERVED_CLIENT_SPOTS,
  STAFF_TRAINING_PERIOD_DAYS,
  SharePointDataService,
  summarizeStaffing
} from './services/staffingDataService';

const MOCK_INTAKE_DATA = [
  {
    id: '1',
    Title: 'Hannah Rodriguez',
    'Intake Status': 'Initial Inquiry',
    Services: 'School Services',
    'Inquiry Date': '2026-07-29',
    'Referral Source': 'School District',
    'School District': 'Denver Public Schools',
    'Start Date': null,
    'Tentative Start Date': '2026-08-15',
    'Staffing Ratio': '1:1',
    DaysToStart: 6,
    BCBA: 'No',
    'Mental Health Services': false,
    'DT Program': false,
    Classroom: false,
    'RMHS Elig': 'No',
    Notes: 'Referred by school psychologist',
    Client: 'Hannah Rodriguez'
  },
  {
    id: '2',
    Title: 'Sterling Chen',
    'Intake Status': 'Contacted Referral',
    Services: 'ABA',
    'Inquiry Date': '2026-06-04',
    'Referral Source': 'Medicaid',
    'School District': 'Aurora Public Schools',
    'Start Date': null,
    'Tentative Start Date': '2026-08-20',
    'Staffing Ratio': '1:1',
    DaysToStart: 11,
    BCBA: 'Yes',
    'Mental Health Services': false,
    'DT Program': false,
    Classroom: false,
    'RMHS Elig': 'No',
    Notes: 'Parent interested in Monday/Wednesday availability',
    Client: 'Sterling Chen'
  },
  {
    id: '3',
    Title: 'Oliver Martinez',
    'Intake Status': 'Sent Intake Email',
    Services: 'ABA',
    'Inquiry Date': '2026-05-12',
    'Referral Source': 'School District',
    'School District': 'Adams 12',
    'Start Date': null,
    'Tentative Start Date': '2026-08-01',
    'Staffing Ratio': '1:1',
    DaysToStart: -9,
    BCBA: 'Yes',
    'Mental Health Services': false,
    'DT Program': false,
    Classroom: true,
    'RMHS Elig': 'Yes',
    Notes: 'Family completed intake',
    Client: 'Oliver Martinez'
  },
  {
    id: '4',
    Title: 'Priya Nair',
    'Intake Status': 'Initial Inquiry',
    Services: 'ABA',
    'Inquiry Date': '2026-08-10',
    'Referral Source': 'Pediatrician',
    'School District': '',
    'Start Date': null,
    'Tentative Start Date': null,
    'Staffing Ratio': '1:1',
    BCBA: 'No',
    'Mental Health Services': false,
    'DT Program': false,
    Classroom: false,
    'RMHS Elig': 'No',
    Notes: 'Just called in, not yet scheduled',
    Client: 'Priya Nair'
  }
];

const MOCK_CURRENT_CLIENTS = [
  {
    id: 'c1',
    Title: 'Sophia Lee',
    Services: 'ABA',
    'Start Date': '2026-06-01',
    'Discharge Date': '2026-09-30',
    'Staffing Ratio': '1:1',
    BCBA: 'Yes',
    'Mental Health Services': false,
    'DT Program': false,
    Classroom: false,
    Status: 'Active'
  },
  {
    id: 'c2',
    Title: 'James Thompson',
    Services: 'ABA',
    'Start Date': '2026-05-15',
    'Discharge Date': '2026-12-31',
    'Staffing Ratio': '1:2',
    BCBA: 'No',
    'Mental Health Services': true,
    'DT Program': true,
    Classroom: true,
    Status: 'Active'
  },
  {
    id: 'c3',
    Title: 'Amelia Brown',
    Services: 'School Services',
    'Start Date': '2026-07-01',
    'Discharge Date': '2026-08-22',
    'Staffing Ratio': '1:1',
    BCBA: 'No',
    'Mental Health Services': false,
    'DT Program': false,
    Classroom: true,
    Status: 'Active'
  }
];

const MOCK_STAFF = [
  { id: 's1', name: 'Marin Patel', role: 'RBT', isActive: true },
  { id: 's2', name: 'Jordan Kim', role: 'RBT', isActive: true },
  { id: 's3', name: 'Avery Nguyen', role: 'BS', isActive: true },
  { id: 's4', name: 'Casey Brooks', role: 'RBT', isActive: true }
];

const INTAKE_STATUS_ORDER = [
  'Initial Inquiry',
  'Contacted Referral',
  'Sent Intake Email',
  'Intake Completed',
  'On Hold'
];

const SERVICES_OPTIONS = ['ABA', 'School Services'];
const STAFFING_RATIOS = ['1:1', '2:1', '1:2'];
const AVAILABILITY_FIELDS = ['MonAM', 'MonPM', 'TueAM', 'TuePM', 'WedAM', 'WedPM', 'ThuAM', 'ThuPM', 'FriAM', 'FriPM'];
const WEEKDAYS = [
  { key: 'Mon', label: 'Monday' },
  { key: 'Tue', label: 'Tuesday' },
  { key: 'Wed', label: 'Wednesday' },
  { key: 'Thu', label: 'Thursday' },
  { key: 'Fri', label: 'Friday' }
];
const MS_PER_DAY = 1000 * 60 * 60 * 24;

function asCalendarDate(value) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  const normalized = new Date(date.getTime());
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}

function dayDifference(startValue, endValue) {
  const startDate = asCalendarDate(startValue);
  const endDate = asCalendarDate(endValue);
  if (!startDate || !endDate) return null;
  return Math.round((endDate.getTime() - startDate.getTime()) / MS_PER_DAY);
}

function normalizeSourceList(value) {
  const normalized = String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  if (normalized.includes('newhire')) return 'newhireonboarding';
  if (normalized.includes('intake')) return 'intake';
  if (normalized.includes('client')) return 'clients';
  if (normalized.includes('staff')) return 'staff';
  return normalized;
}

function normalizeIntakeStatus(value) {
  const normalized = String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const statusMap = {
    initialinquiry: 'Initial Inquiry',
    contactedreferral: 'Contacted Referral',
    sentintakeemail: 'Sent Intake Email',
    intakecompleted: 'Intake Completed',
    readytostartintakecompleted: 'Intake Completed',
    onhold: 'On Hold'
  };
  return statusMap[normalized] || String(value || '').trim();
}

function matchesServiceAndRatio(item, filters) {
  if (filters.service && item.Services !== filters.service) return false;
  if (filters.staffingRatio && item['Staffing Ratio'] !== filters.staffingRatio) return false;
  return true;
}

// Intake Status values are funnel-stage labels (Initial Inquiry, Enrolled, ...) that don't
// meaningfully apply to already-active current clients, so only intake/prospect data is
// filtered on it - see matchesServiceAndRatio for the current-clients-safe subset.
function matchesFilters(item, filters) {
  if (!matchesServiceAndRatio(item, filters)) return false;
  if (filters.intakeStatus && normalizeIntakeStatus(item['Intake Status']) !== filters.intakeStatus) return false;
  return true;
}

function alphabetizeNames(items, getName = (item) => item.name || item.Title) {
  return [...items].sort((left, right) =>
    String(getName(left) || '').localeCompare(String(getName(right) || ''), undefined, { sensitivity: 'base' })
  );
}

export default function IntakeCapacityPlanner() {
  const [authToken, setAuthToken] = useState(null);
  const [intakeData, setIntakeData] = useState(MOCK_INTAKE_DATA);
  const [currentClients, setCurrentClients] = useState(MOCK_CURRENT_CLIENTS);
  const [staffingSnapshot, setStaffingSnapshot] = useState(() => summarizeStaffing(MOCK_STAFF));
  const [prospectiveStaff, setProspectiveStaff] = useState([]);
  const [capacityPlanning, setCapacityPlanning] = useState([]);
  const [syncLoading, setSyncLoading] = useState(false);
  const [syncConnected, setSyncConnected] = useState(false);
  const [syncError, setSyncError] = useState('');
  const [lastSyncedAt, setLastSyncedAt] = useState(null);
  const [view, setView] = useState('dashboard');
  const [projectionWeeks, setProjectionWeeks] = useState(8);
  const [historyWeeks, setHistoryWeeks] = useState(2);
  const [availabilityProjectionWeeks, setAvailabilityProjectionWeeks] = useState(0);
  const [whatIfStaffDelta, setWhatIfStaffDelta] = useState(0);
  const [whatIfClientDelta, setWhatIfClientDelta] = useState(0);
  const [filters, setFilters] = useState({
    service: null,
    staffingRatio: null,
    intakeStatus: null,
    showProspectsOnly: true
  });
  const [expandedRow, setExpandedRow] = useState(null);
  const [sharePointConfig] = useState(() => getSharePointConfig());
  const [sharePointService] = useState(() => new SharePointDataService(sharePointConfig));

  useEffect(() => {
    setAuthToken('mock-token');
  }, []);

  // Redirect login reloads the page; resume the SharePoint sync automatically if a session already exists.
  const hasCheckedSessionRef = useRef(false);
  useEffect(() => {
    // Guard against React.StrictMode's double-invoke in dev, which would otherwise
    // fire two concurrent sign-in/sync attempts on every reload.
    if (hasCheckedSessionRef.current) return;
    hasCheckedSessionRef.current = true;

    sharePointService
      .hasSession()
      .then((hasSession) => {
        if (hasSession) connectSharePoint();
      })
      .catch(() => {});
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleSignIn = async () => {
    setSyncError('');
    try {
      await sharePointService.login();
      // loginRedirect navigates away; execution won't reach here until the app reloads.
    } catch (error) {
      console.error('SharePoint sign-in failed:', error);
      setSyncError(error.errorCode ? `${error.errorCode}: ${error.errorMessage || error.message}` : error.message || 'Failed to sign in');
    }
  };

  const connectSharePoint = async () => {
    setSyncLoading(true);
    setSyncError('');

    try {
      if (!sharePointService.isConfigured()) {
        throw new Error('Missing VITE_SP_CLIENT_ID or VITE_SP_TENANT_ID in Intake .env.local');
      }

      const snapshot = await sharePointService.loadSnapshot();
      setStaffingSnapshot(snapshot.staffing);
      setCurrentClients(snapshot.currentClients);
      setIntakeData(snapshot.intakeClients);
      setProspectiveStaff(snapshot.prospectiveStaff || []);
      setCapacityPlanning(snapshot.capacityPlanning || []);
      setSyncConnected(true);
      setLastSyncedAt(snapshot.lastUpdatedAt);
    } catch (error) {
      // Log the raw MSAL/SharePoint error (errorCode, errorMessage) for diagnosis - the UI only shows a short message.
      console.error('SharePoint connect failed:', error);
      setSyncConnected(false);
      setSyncError(error.errorCode ? `${error.errorCode}: ${error.errorMessage || error.message}` : error.message || 'Failed to load SharePoint data');
    } finally {
      setSyncLoading(false);
    }
  };

  const filteredIntakeData = intakeData.filter((item) => matchesFilters(item, filters));
  const filteredCurrentClients = currentClients.filter((item) => matchesServiceAndRatio(item, filters));
  const sortedIntakeData = alphabetizeNames(filteredIntakeData);
  const sortedCurrentClients = alphabetizeNames(filteredCurrentClients);

  const funnelData = INTAKE_STATUS_ORDER.map((status) => ({
    status,
    count: filteredIntakeData.filter((item) => normalizeIntakeStatus(item['Intake Status']) === status).length
  }));

  const now = asCalendarDate(new Date());
  const availabilityProjectionDate = new Date(now);
  availabilityProjectionDate.setDate(availabilityProjectionDate.getDate() + availabilityProjectionWeeks * 7);
  const isLiveByAvailabilityDate = (person) => {
    if (!person.startDate) return false;
    const liveDate = new Date(person.startDate);
    liveDate.setDate(liveDate.getDate() + STAFF_TRAINING_PERIOD_DAYS);
    return liveDate <= availabilityProjectionDate;
  };
  const projectedAvailabilityStaff = staffingSnapshot.activeCount + prospectiveStaff.filter(
    (person) =>
      isAvailableDirectStaffRole(person.position) &&
      isLiveByAvailabilityDate(person)
  ).length;
  const projectedAvailabilityTraining = prospectiveStaff.filter(
    (person) =>
      isAvailableDirectStaffRole(person.position) &&
      person.startDate &&
      !isLiveByAvailabilityDate(person) &&
      new Date(person.startDate) <= availabilityProjectionDate
  ).length;

  // Every prospect with a recorded Inquiry Date gets a bar. If they have a future start date,
  // the bar should extend to that date; otherwise it remains open-ended and runs to today.
  const validTimelines = filteredIntakeData
    .map((item) => {
      const inquiryValue = item['Inquiry Date'];
      if (!inquiryValue) return null;

      const inquiryDate = asCalendarDate(inquiryValue);
      if (!inquiryDate) return null;

      const startValue = item['Start Date'] || item['Tentative Start Date'];
      const startDate = startValue ? asCalendarDate(startValue) : null;
      const isOpenEnded = !startDate;
      const endDate = startDate || now;
      const endDateLabel = startDate ? formatDate(startValue) : 'today';

      return {
        id: item.id,
        name: item.Title,
        inquiryTimestamp: inquiryDate.getTime(),
        spanMs: Math.max(0, endDate.getTime() - inquiryDate.getTime()),
        daysToStart: Math.round((endDate.getTime() - inquiryDate.getTime()) / MS_PER_DAY),
        inquiryDateLabel: formatDate(inquiryValue),
        startDateLabel: startDate ? formatDate(startValue) : null,
        endDateLabel,
        isTentative: !item['Start Date'] && Boolean(item['Tentative Start Date']),
        isOpenEnded
      };
    })
    .filter(Boolean)
    .sort((left, right) => right.inquiryTimestamp - left.inquiryTimestamp);

  const timelineMinDate = validTimelines.length
    ? Math.min(...validTimelines.map((row) => row.inquiryTimestamp)) - 14 * MS_PER_DAY
    : null;

  const timelineMaxDate = validTimelines.length
    ? Math.max(...validTimelines.map((row) => row.inquiryTimestamp + row.spanMs)) + 14 * MS_PER_DAY
    : null;

  const inquiryToStartTimeline = alphabetizeNames(validTimelines.map((row) => ({
    ...row,
    leftPercent: timelineMinDate !== null && timelineMaxDate !== null
      ? ((row.inquiryTimestamp - timelineMinDate) / (timelineMaxDate - timelineMinDate)) * 100
      : 0,
    widthPercent: timelineMinDate !== null && timelineMaxDate !== null
      ? (row.spanMs / (timelineMaxDate - timelineMinDate)) * 100
      : 0
  })), (row) => row.name);

  const timelineRangeMs = timelineMinDate !== null && timelineMaxDate !== null ? timelineMaxDate - timelineMinDate : 0;
  const todayPositionPct = timelineRangeMs > 0 ? ((now.getTime() - timelineMinDate) / timelineRangeMs) * 100 : 0;

  const projectionData = applyWhatIfAdjustments(
    generateCapacityProjection(
      filteredCurrentClients,
      filteredIntakeData,
      staffingSnapshot.activeCount,
      prospectiveStaff,
      projectionWeeks,
      staffingSnapshot.staff,
      historyWeeks
    ),
    { staffDelta: whatIfStaffDelta, clientDelta: whatIfClientDelta }
  );

  // Highlights the weeks where required staff outpaces projected staff, as a fill under the lines.
  const projectionChartData = projectionData.map((week) => {
    const staffingGap = Number((week.projectedStaff - week.requiredStaff).toFixed(2));
    return {
      ...week,
      staffingGap,
      staffAboveRequired: Math.max(0, staffingGap),
      staffBelowRequired: Math.min(0, staffingGap),
      shortfall: Math.max(0, week.requiredStaff - week.projectedStaff)
    };
  });

  const historicalProjectionWeeks = projectionChartData.filter((week) => week.isHistorical);
  const currentProjectionWeek = projectionChartData.find((week) => !week.isHistorical);

  const upcomingStaffEndDates = staffingSnapshot.staff
    .filter((person) => person.endDate)
    .map((person) => ({
      ...person,
      daysUntilEnd: dayDifference(new Date(), person.endDate)
    }))
    .filter((person) => person.daysUntilEnd !== null && person.daysUntilEnd >= 0)
    .sort((left, right) => String(left.name || '').localeCompare(String(right.name || ''), undefined, { sensitivity: 'base' }));

  const planningRows = alphabetizeNames([
    ...currentClients
      .filter((client) => client['Start Date'] || client['Tentative Start Date'])
      .map((client) => ({
      id: `client-${client.id}`,
      sourceList: 'Clients',
      sourceId: String(client.id),
      personType: 'Client',
      name: client.Title,
      position: client.Services,
      isUpcoming: false
      })),
    ...intakeData
      .filter((client) => client['Start Date'] || client['Tentative Start Date'])
      .map((client) => ({
      id: `intake-${client.id}`,
      sourceList: 'Intake',
      sourceId: String(client.id),
      personType: 'Client',
      name: client.Title,
      position: client.Services,
      isUpcoming: true
      }))
  ].map((person) => ({
    ...person,
    planning: capacityPlanning.find(
      (row) =>
        normalizeSourceList(row.sourceList) === normalizeSourceList(person.sourceList) &&
        row.sourceId === person.sourceId
    )
  })), (person) => person.name);

  const availabilityTotals = WEEKDAYS.map((day) => ({
    ...day,
    amClients: planningRows.filter((person) => person.planning?.availability?.[`${day.key}AM`]).length,
    pmClients: planningRows.filter((person) => person.planning?.availability?.[`${day.key}PM`]).length,
    amRequiredStaff: Number((planningRows.reduce((sum, person) => sum + (person.planning?.availability?.[`${day.key}AM`] ? getRequiredStaffUnits(person) : 0), RESERVED_CLIENT_SPOTS)).toFixed(2)),
    pmRequiredStaff: Number((planningRows.reduce((sum, person) => sum + (person.planning?.availability?.[`${day.key}PM`] ? getRequiredStaffUnits(person) : 0), RESERVED_CLIENT_SPOTS)).toFixed(2)),
    amAvailableStaff: projectedAvailabilityStaff,
    pmAvailableStaff: projectedAvailabilityStaff
  }));

  const availabilityGaps = availabilityTotals.flatMap((day) => [
    {
      key: `${day.key}-AM`,
      label: `${day.label} AM`,
      clients: day.amClients,
      requiredStaff: day.amRequiredStaff,
      availableStaff: day.amAvailableStaff
    },
    {
      key: `${day.key}-PM`,
      label: `${day.label} PM`,
      clients: day.pmClients,
      requiredStaff: day.pmRequiredStaff,
      availableStaff: day.pmAvailableStaff
    }
  ]).map((period) => {
    const gap = Number((period.availableStaff - period.requiredStaff).toFixed(2));
    return {
      ...period,
      gap,
      additionalClientUnits: Math.max(0, Math.floor((gap + 0.0001) / REQUIRED_STAFF_PER_CLIENT)),
      status: gap > 0 ? 'Open capacity' : gap < 0 ? 'Shortfall' : 'At capacity'
    };
  });

  const updateAvailability = async (planningRow, field, value) => {
    let planning = planningRow.planning;
    if (!planning) {
      const availability = Object.fromEntries(AVAILABILITY_FIELDS.map((name) => [name, false]));
      availability[field] = value;
      const created = await sharePointService.createListItem(
        sharePointConfig.capacityPlanningSiteUrl,
        sharePointConfig.capacityPlanningListName,
        {
          Title: planningRow.name,
          SourceList: planningRow.sourceList,
          SourceID: Number(planningRow.sourceId),
          PersonType: planningRow.personType,
          ...availability
        }
      );
      planning = {
        id: created.Id,
        sourceList: planningRow.sourceList,
        sourceId: planningRow.sourceId,
        personType: planningRow.personType,
        availability
      };
    } else {
      await sharePointService.updateListItem(
        sharePointConfig.capacityPlanningSiteUrl,
        sharePointConfig.capacityPlanningListName,
        planning.id,
        { [field]: value }
      );
      planning = { ...planning, availability: { ...planning.availability, [field]: value } };
    }

    setCapacityPlanning((rows) => {
      const existing = rows.some((row) => row.id === planning.id);
      return existing ? rows.map((row) => (row.id === planning.id ? planning : row)) : [...rows, planning];
    });
  };

  const setAllRowAvailability = async (planningRow, value) => {
    const availability = Object.fromEntries(AVAILABILITY_FIELDS.map((field) => [field, value]));
    let planning = planningRow.planning;
    if (!planning) {
      const created = await sharePointService.createListItem(sharePointConfig.capacityPlanningSiteUrl, sharePointConfig.capacityPlanningListName, {
        Title: planningRow.name,
        SourceList: planningRow.sourceList,
        SourceID: Number(planningRow.sourceId),
        PersonType: planningRow.personType,
        ...availability
      });
      planning = { id: created.Id, sourceList: planningRow.sourceList, sourceId: planningRow.sourceId, personType: planningRow.personType };
      setCapacityPlanning((rows) => [...rows, { ...planning, availability }]);
    } else {
      await sharePointService.updateListItem(sharePointConfig.capacityPlanningSiteUrl, sharePointConfig.capacityPlanningListName, planning.id, availability);
      setCapacityPlanning((rows) => rows.map((row) => row.id === planning.id ? { ...row, availability } : row));
    }
  };

  const stats = {
    totalProspects: filteredIntakeData.length,
    readyToEnroll: filteredIntakeData.filter(
      (item) => item['Intake Status'] === 'Enrolled' || (item.DaysToStart && item.DaysToStart <= 7)
    ).length,
    byService: Object.fromEntries(
      SERVICES_OPTIONS.map((service) => [service, filteredIntakeData.filter((item) => item.Services === service).length])
    ),
    currentCapacity: filteredCurrentClients.length,
    totalClients: filteredCurrentClients.length + filteredIntakeData.length,
    averageStaffingRatio: calculateAverageRatio([...filteredIntakeData, ...filteredCurrentClients]),
    activeStaff: staffingSnapshot.activeCount
  };

  const staffRows = alphabetizeNames([
    ...staffingSnapshot.staff.map((person) => ({
      id: `current-${person.id}`,
      name: person.name,
      position: person.role,
      source: 'Staff list',
      status: 'Active',
      startDate: person.startDate,
      endDate: person.endDate
    })),
    ...prospectiveStaff.map((person) => ({
      id: `upcoming-${person.id}`,
      name: person.name,
      position: person.position,
      source: 'HR New Hire Onboarding',
      status: person.status || 'Upcoming',
      startDate: person.startDate,
      endDate: person.endDate || null
    }))
  ]);

  return (
    <div className="min-h-screen bg-[#f1f1f1]">
      <header className="sticky top-0 z-40 border-b border-[#352b43] bg-[#3f425e] shadow-sm">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
          <div>
            <h1 className="text-3xl font-bold text-white">Intake Capacity Planner</h1>
            <p className="mt-1 text-sm text-slate-300">Track prospective clients and project staffing needs</p>
          </div>

          {authToken ? (
            <div className="flex gap-3">
              <button
                onClick={connectSharePoint}
                disabled={syncLoading}
                className="flex items-center gap-2 rounded-lg bg-[#edf3f2] px-4 py-2 font-medium text-[#352b43] transition-colors hover:bg-[#dce9e7] disabled:cursor-not-allowed disabled:opacity-60"
              >
                {syncLoading ? 'Syncing...' : syncConnected ? 'Refresh SharePoint Data' : 'Connect SharePoint'}
              </button>
              <button
                onClick={() => exportIntakeDataToCSV(filteredIntakeData)}
                className="flex items-center gap-2 rounded-lg bg-[#352b43] px-4 py-2 font-semibold text-white transition-colors hover:bg-[#211b21] focus:outline-none focus:ring-2 focus:ring-[#d9a441]"
              >
                <Download size={18} />
                Export
              </button>
              <select
                value={view}
                onChange={(event) => setView(event.target.value)}
                className="rounded-lg border border-[#b8b5b3] bg-white px-4 py-2 font-medium text-[#3f425e] hover:border-[#3f425e] focus:outline-none focus:ring-2 focus:ring-[#352b43]"
              >
                <option value="dashboard">Dashboard</option>
                <option value="details">Prospective Clients</option>
                <option value="currentClients">Current Clients</option>
                <option value="staff">Staff</option>
                <option value="availability">Availability Planning</option>
                <option value="projections">Capacity Projections</option>
              </select>
            </div>
          ) : (
            <button
              onClick={handleSignIn}
              className="rounded-lg bg-[#d9a441] px-6 py-2 font-medium text-[#211b21] transition-colors hover:bg-[#e2b35a]"
            >
              Sign in with Microsoft
            </button>
          )}
        </div>
      </header>

      <main className="mx-auto max-w-7xl px-6 py-8">
        <div className="mb-4 rounded-lg border border-slate-200 bg-white p-4 text-sm text-slate-700 shadow-sm">
          <p>
            Data source: {syncConnected ? 'SharePoint (Staff, Current Clients, Intake, HR New Hire Onboarding)' : 'Local mock snapshot'}
            {lastSyncedAt ? ` • updated ${new Date(lastSyncedAt).toLocaleTimeString()}` : ''}
          </p>
          {syncError && <p className="mt-1 text-[#a84b2a]">{syncError}</p>}
        </div>

        <div className="mb-6 rounded-lg border border-slate-200 bg-white p-4 shadow-sm">
          <div className="mb-4 flex items-center gap-2">
            <Filter size={18} className="text-slate-600" />
            <h3 className="font-semibold text-slate-900">Filters</h3>
          </div>
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
            <FilterSelect
              label="Service"
              options={SERVICES_OPTIONS}
              value={filters.service}
              onChange={(value) => setFilters({ ...filters, service: value })}
            />
            <FilterSelect
              label="Staffing Ratio"
              options={STAFFING_RATIOS}
              value={filters.staffingRatio}
              onChange={(value) => setFilters({ ...filters, staffingRatio: value })}
            />
            <FilterSelect
              label="Intake Status"
              options={INTAKE_STATUS_ORDER}
              value={filters.intakeStatus}
              onChange={(value) => setFilters({ ...filters, intakeStatus: value })}
            />
            <button
              onClick={() =>
                setFilters({
                  service: null,
                  staffingRatio: null,
                  intakeStatus: null,
                  showProspectsOnly: true
                })
              }
              className="rounded-lg bg-slate-100 px-4 py-2 font-medium text-slate-700 transition-colors hover:bg-slate-200"
            >
              Clear Filters
            </button>
          </div>
          {view === 'projections' && (
            <p className="mt-3 text-xs text-slate-500">
              Service and Staffing Ratio filter both current clients and the intake pipeline below. Intake Status
              only applies to the pipeline, since active current clients aren't staged in the intake funnel.
            </p>
          )}
        </div>

        {view === 'dashboard' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2 lg:grid-cols-4">
              <StatCard title="Total Clients" value={stats.totalClients} subtitle="current + incoming intake" color="blue" />
              <StatCard title="Current Clients" value={stats.currentCapacity} subtitle="active enrollments" color="purple" />
              <StatCard title="Prospective Clients" value={stats.totalProspects} subtitle="in intake pipeline" color="green" />
              <StatCard title="Available Staff" value={stats.activeStaff} subtitle="active RBT, BT1, and BT2" color="orange" />
            </div>

            <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
              <h2 className="mb-4 text-xl font-bold text-slate-900">Intake Funnel</h2>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={funnelData} layout="vertical" margin={{ left: 150, right: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                  <XAxis type="number" stroke="#64748b" />
                  <YAxis dataKey="status" type="category" width={140} tick={{ fontSize: 12 }} stroke="#64748b" />
                  <Tooltip
                    contentStyle={{ backgroundColor: '#fff', border: '1px solid #e2e8f0', borderRadius: '8px' }}
                    formatter={(value) => [value, 'Count']}
                  />
                  <Bar dataKey="count" fill="#3f425e" radius={[0, 8, 8, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>

            <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
              <h2 className="mb-1 text-xl font-bold text-slate-900">Inquiry-to-Start Timeline</h2>
              <p className="mb-4 text-sm text-slate-600">
                Every prospect with a recorded inquiry date, from first contact to their start date. Dark green =
                confirmed start date, light green = tentative start date, gray = no start date set yet (bar runs to
                today). The vertical marker shows today.
              </p>
              {inquiryToStartTimeline.length === 0 ? (
                <p className="text-sm text-slate-500">No intake records have an Inquiry Date yet.</p>
              ) : (
                <div className="max-h-[600px] overflow-x-auto overflow-y-auto pr-2">
                  <div className="relative min-h-[220px] min-w-[900px]">
                    <div className="pointer-events-none absolute inset-y-0 left-[192px] right-[60px] z-20">
                      <div
                        className="absolute inset-y-0 w-px bg-slate-800/70"
                        style={{ left: `${todayPositionPct}%` }}
                      >
                        <span className="absolute left-1 top-0 -translate-y-full whitespace-nowrap text-[10px] font-semibold uppercase tracking-wide text-slate-600">
                          Today
                        </span>
                      </div>
                    </div>

                    {inquiryToStartTimeline.map((row) => {
                      const leftPercent = row.leftPercent;
                      const widthPercent = row.widthPercent;
                      const barColor = row.isOpenEnded ? '#b8b5b3' : row.isTentative ? '#a7c0be' : '#352b43';

                      return (
                        <div key={row.id} className="mb-3 flex items-center gap-3">
                          <div className="w-[180px] truncate text-right text-sm font-medium text-slate-700">
                            {row.name}
                          </div>

                          <div className="relative h-6 flex-1 overflow-hidden rounded bg-slate-100">
                            {row.isOpenEnded && (
                              <span
                                className="absolute inset-y-0 -translate-x-full pr-2 text-right text-xs font-medium leading-6 text-slate-500"
                                style={{ left: `${leftPercent}%` }}
                              >
                                {row.inquiryDateLabel}
                              </span>
                            )}
                            <div
                              className="absolute inset-y-0 rounded"
                              style={{
                                left: `${leftPercent}%`,
                                width: `${Math.max(widthPercent, 2)}%`,
                                backgroundColor: barColor,
                                border: row.isOpenEnded
                                  ? '1px solid rgba(184,181,179,0.9)'
                                  : row.isTentative
                                    ? '1px solid rgba(167,192,190,0.95)'
                                    : '1px solid rgba(53,43,67,0.85)'
                              }}
                            >
                              {row.isTentative && (
                                <span
                                  className="relative z-30 block whitespace-nowrap overflow-visible rounded px-2 text-[11px] font-semibold leading-6 text-[#352b43]"
                                  title={`Tentative Start Date: ${row.startDateLabel}`}
                                >
                                  Tentative Start Date: {row.startDateLabel}
                                </span>
                              )}
                            </div>
                          </div>

                          <div className="w-12 text-right text-xs font-medium text-slate-600">{row.daysToStart}d</div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 gap-6 lg:grid-cols-2">
              <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <h2 className="mb-4 text-xl font-bold text-slate-900">Prospective Clients by Service</h2>
                <ResponsiveContainer width="100%" height={250}>
                  <PieChart>
                    <Pie
                      data={Object.entries(stats.byService).map(([name, value]) => ({ name, value }))}
                      cx="50%"
                      cy="50%"
                      labelLine={false}
                      label={({ name, value }) => `${name}: ${value}`}
                      outerRadius={80}
                      fill="#352b43"
                      dataKey="value"
                    >
                      <Cell fill="#3f425e" />
                      <Cell fill="#a7c0be" />
                      <Cell fill="#d9a441" />
                      <Cell fill="#a84b2a" />
                      <Cell fill="#352b43" />
                    </Pie>
                    <Tooltip />
                  </PieChart>
                </ResponsiveContainer>
              </div>

              <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <h2 className="mb-4 flex items-center gap-2 text-xl font-bold text-slate-900">
                  <AlertCircle size={20} className="text-[#d9a441]" />
                  Upcoming Enrollments
                </h2>
                <div className="space-y-3">
                  {sortedIntakeData
                    .map((item) => ({
                      item,
                      daysUntilStart: dayDifference(new Date(), item['Start Date'] || item['Tentative Start Date'])
                    }))
                    .filter(({ daysUntilStart }) => daysUntilStart !== null && daysUntilStart >= 0 && daysUntilStart <= 14)
                    .sort((left, right) => left.daysUntilStart - right.daysUntilStart)
                    .slice(0, 5)
                    .map(({ item, daysUntilStart }, index) => (
                      <div key={item.id ?? index} className="flex items-start justify-between rounded-lg border-2 border-[#a84b2a] bg-[#a7c0be]/35 p-3">
                        <div>
                          <p className="font-medium text-slate-900">{item.Title}</p>
                          <p className="text-sm text-slate-600">{item.Services}</p>
                        </div>
                        <div className="text-right">
                          <p className="font-bold text-[#a84b2a]">{daysUntilStart}d</p>
                          <p className="text-xs text-slate-600">until start</p>
                        </div>
                      </div>
                    ))}
                </div>
              </div>
            </div>
          </div>
        )}

        {view === 'details' && (
          <div className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="border-b border-slate-200 bg-slate-50">
                  <tr>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Client</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Service</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Inquiry Date</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Status</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Start Date</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Staffing</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700"></th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {sortedIntakeData.map((item, index) => (
                    <React.Fragment key={item.id ?? index}>
                      <tr
                        className="cursor-pointer transition-colors hover:bg-slate-50"
                        onClick={() => setExpandedRow(expandedRow === item.id ? null : item.id)}
                      >
                        <td className="px-6 py-4 text-sm font-medium text-slate-900">{item.Title}</td>
                        <td className="px-6 py-4 text-sm text-slate-600">{item.Services}</td>
                        <td className="px-6 py-4 text-sm text-slate-600">{formatDate(item['Inquiry Date'])}</td>
                        <td className="px-6 py-4">
                          <StatusBadge status={item['Intake Status']} />
                        </td>
                        <td className="px-6 py-4 text-sm text-slate-600">{formatDate(item['Start Date'] || item['Tentative Start Date'])}</td>
                        <td className="px-6 py-4 text-sm font-medium text-slate-900">{item['Staffing Ratio']}</td>
                        <td className="px-6 py-4 text-right">
                          <ChevronDown
                            size={18}
                            className={`text-slate-400 transition-transform ${expandedRow === item.id ? 'rotate-180' : ''}`}
                          />
                        </td>
                      </tr>
                      {expandedRow === item.id && (
                        <tr className="bg-slate-50">
                          <td colSpan="7" className="px-6 py-4">
                            <DetailPanel item={item} />
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === 'currentClients' && (
          <div className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="border-b border-slate-200 bg-slate-50">
                  <tr>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Client</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Program</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Service</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Start Date</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">DC Date</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Staffing</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Status</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {sortedCurrentClients.map((client, index) => (
                    <tr key={client.id ?? index} className="hover:bg-slate-50">
                      <td className="px-6 py-4 text-sm font-medium text-slate-900">{client.Title}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{client.Program || 'N/A'}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{client.Services}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{formatDate(client['Start Date']) || 'N/A'}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{formatDate(client['Discharge Date']) || 'N/A'}</td>
                      <td className="px-6 py-4 text-sm font-medium text-slate-900">{client['Staffing Ratio']}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{client.Status}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === 'staff' && (
          <div className="overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="border-b border-slate-200 bg-slate-50">
                  <tr>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Staff Member</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Position</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Source</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Status</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">Start Date</th>
                    <th className="px-6 py-3 text-left text-xs font-semibold text-slate-700">End Date</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {staffRows.map((person, index) => (
                    <tr key={person.id ?? index} className="hover:bg-slate-50">
                      <td className="px-6 py-4 text-sm font-medium text-slate-900">{person.name}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{person.position || 'N/A'}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{person.source}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{person.status}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{formatDate(person.startDate) || 'N/A'}</td>
                      <td className="px-6 py-4 text-sm text-slate-600">{formatDate(person.endDate) || 'N/A'}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {view === 'availability' && (
          <div className="space-y-6">
            <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
              <div className="mb-2 flex flex-wrap items-center justify-between gap-3">
                <h2 className="text-xl font-bold text-slate-900">Availability Planning</h2>
                <label className="flex items-center gap-2 text-sm font-medium text-slate-700">
                  <span>Project staffing for</span>
                  <select
                    value={availabilityProjectionWeeks}
                    onChange={(event) => setAvailabilityProjectionWeeks(Number(event.target.value))}
                    className="rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm font-medium text-slate-700"
                  >
                    <option value={0}>Today</option>
                    <option value={1}>1 week out</option>
                    <option value={2}>2 weeks out</option>
                    <option value={4}>4 weeks out</option>
                    <option value={8}>8 weeks out</option>
                  </select>
                </label>
              </div>
              <p className="mb-4 text-sm text-slate-600">
                Check each AM/PM period from the Capacity Planning SharePoint list. Required staff is based on the clients present in each period. The projection includes {projectedAvailabilityStaff} live direct-care staff by {formatDate(availabilityProjectionDate)}; {projectedAvailabilityTraining} additional hire{projectedAvailabilityTraining === 1 ? '' : 's'} {projectedAvailabilityTraining === 1 ? 'is' : 'are'} still in training.
              </p>
              <div className="mb-6 grid grid-cols-1 gap-3 md:grid-cols-5">
                {availabilityTotals.map((day) => (
                  <div key={`top-${day.key}`} className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                    <p className="font-semibold text-slate-900">{day.label}</p>
                    <p className="text-sm text-slate-600">AM: <span className="font-bold text-slate-900">{day.amClients}</span> clients / <span className="font-bold text-slate-900">{day.amRequiredStaff}</span> staff</p>
                    <p className="text-sm text-slate-600">PM: <span className="font-bold text-slate-900">{day.pmClients}</span> clients / <span className="font-bold text-slate-900">{day.pmRequiredStaff}</span> staff</p>
                  </div>
                ))}
              </div>
              <div className="mb-6 rounded-lg border border-slate-200 bg-[#f8faf9] p-5">
                <div className="mb-4 flex flex-wrap items-end justify-between gap-3">
                  <div>
                    <h3 className="text-lg font-bold text-slate-900">Open Schedule Gaps</h3>
                    <p className="mt-1 text-sm text-slate-600">
                      Each period includes {RESERVED_CLIENT_SPOTS} reserved client spot plus {REQUIRED_STAFF_PER_CLIENT} staff units per scheduled client.
                    </p>
                  </div>
                  <div className="flex gap-3 text-xs font-medium text-slate-600">
                    <span><span className="mr-1 inline-block h-2 w-2 rounded-full bg-[#78a995]" />Open</span>
                    <span><span className="mr-1 inline-block h-2 w-2 rounded-full bg-[#d9a441]" />Full</span>
                    <span><span className="mr-1 inline-block h-2 w-2 rounded-full bg-[#b45b46]" />Shortfall</span>
                  </div>
                </div>
                <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-5">
                  {availabilityGaps.map((period) => {
                    const isOpen = period.gap > 0;
                    const isShort = period.gap < 0;
                    const statusColor = isOpen
                      ? 'border-[#a7c8bb] bg-[#edf6f1]'
                      : isShort
                        ? 'border-[#e0b5a8] bg-[#fbefeb]'
                        : 'border-[#e6ca83] bg-[#fff8e5]';
                    const statusTextColor = isOpen
                      ? 'text-[#34735d]'
                      : isShort
                        ? 'text-[#a84b2a]'
                        : 'text-[#9a6b08]';

                    return (
                      <div key={period.key} className={`rounded-lg border p-3 ${statusColor}`}>
                        <div className="flex items-center justify-between gap-2">
                          <p className="font-semibold text-slate-900">{period.label}</p>
                          <span className={`text-xs font-bold ${statusTextColor}`}>{period.status}</span>
                        </div>
                        <div className="mt-3 flex items-end justify-between gap-2">
                          <div>
                            <p className="text-2xl font-bold text-slate-900">{Math.abs(period.gap)}</p>
                            <p className="text-xs text-slate-600">{isOpen ? 'staff units open' : isShort ? 'staff units short' : 'staff units remaining'}</p>
                          </div>
                          <p className="text-right text-xs text-slate-600">
                            {period.clients} clients<br />
                            {period.requiredStaff} / {period.availableStaff} staff
                          </p>
                        </div>
                        <p className={`mt-3 border-t pt-2 text-xs font-semibold ${statusTextColor} ${isOpen ? 'border-[#c9dfd5]' : isShort ? 'border-[#edccc2]' : 'border-[#f0dfa9]'}`}>
                          {isOpen ? `Fits ${period.additionalClientUnits} more client${period.additionalClientUnits === 1 ? '' : 's'}` : 'No additional client capacity'}
                        </p>
                      </div>
                    );
                  })}
                </div>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full min-w-[950px]">
                  <thead className="border-b border-slate-200 bg-slate-50">
                    <tr>
                      <th className="px-4 py-3 text-left text-xs font-semibold text-slate-700">Name</th>
                      <th className="px-4 py-3 text-left text-xs font-semibold text-slate-700">Service</th>
                      <th className="px-2 py-3 text-center text-xs font-semibold text-slate-700">Row</th>
                      {AVAILABILITY_FIELDS.map((field) => (
                        <th key={field} className="px-2 py-3 text-center text-xs font-semibold text-slate-700">{field}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {planningRows.map((person) => (
                      <tr key={person.id} className={person.isUpcoming ? 'bg-[#fff8e5] hover:bg-[#fff3cc]' : 'hover:bg-slate-50'}>
                        <td className="px-4 py-3 text-sm font-medium text-slate-900">
                          <div className="flex items-center gap-2">
                            <span>{person.name}</span>
                            {person.isUpcoming && (
                              <span className="rounded-full bg-[#d9a441]/25 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-[#9a6b08]">
                                Upcoming
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="px-4 py-3 text-sm text-slate-600">{person.position || 'N/A'}</td>
                        <td className="px-2 py-3 text-center">
                          <button
                            type="button"
                            onClick={() => setAllRowAvailability(person, true)}
                            className="text-xs font-medium text-[#3f425e] hover:text-[#352b43]"
                          >
                            Check all
                          </button>
                          <button
                            type="button"
                            onClick={() => setAllRowAvailability(person, false)}
                            className="ml-2 text-xs font-medium text-slate-500 hover:text-slate-700"
                          >
                            Clear
                          </button>
                        </td>
                        {AVAILABILITY_FIELDS.map((field) => (
                          <td key={field} className="px-2 py-3 text-center">
                            <input
                              type="checkbox"
                              checked={Boolean(person.planning?.availability?.[field])}
                              title={`Save ${field} for ${person.name}`}
                              onChange={(event) => updateAvailability(person, field, event.target.checked)}
                              className="h-4 w-4 accent-[#3f425e] disabled:cursor-not-allowed"
                            />
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="mt-6 grid grid-cols-1 gap-3 border-t border-slate-200 pt-6 md:grid-cols-5">
                {availabilityTotals.map((day) => (
                  <div key={`bottom-${day.key}`} className="rounded-lg border border-slate-200 bg-slate-50 p-3">
                    <p className="font-semibold text-slate-900">{day.label}</p>
                    <p className="text-sm text-slate-600">AM: <span className="font-bold text-slate-900">{day.amClients}</span> clients / <span className="font-bold text-slate-900">{day.amRequiredStaff}</span> staff</p>
                    <p className="text-sm text-slate-600">PM: <span className="font-bold text-slate-900">{day.pmClients}</span> clients / <span className="font-bold text-slate-900">{day.pmRequiredStaff}</span> staff</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {view === 'projections' && (
          <div className="space-y-6">
            <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
              <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
                <h2 className="text-xl font-bold text-slate-900">Projection Controls</h2>
                <button
                  type="button"
                  onClick={() => {
                    setWhatIfStaffDelta(0);
                    setWhatIfClientDelta(0);
                  }}
                  className="rounded-lg bg-slate-100 px-4 py-2 text-sm font-medium text-slate-700 transition-colors hover:bg-slate-200"
                >
                  Reset scenario
                </button>
              </div>
              <div className="grid grid-cols-1 gap-6 md:grid-cols-2 lg:grid-cols-4">
                <div>
                  <label className="mb-2 flex items-center justify-between text-sm font-medium text-slate-700">
                    <span>Weeks ahead</span>
                    <span className="font-bold text-slate-900">{projectionWeeks}</span>
                  </label>
                  <input
                    type="range"
                    min={4}
                    max={26}
                    step={1}
                    value={projectionWeeks}
                    onChange={(event) => setProjectionWeeks(Number(event.target.value))}
                    className="w-full accent-[#3f425e]"
                  />
                </div>
                <div>
                  <label className="mb-2 flex items-center justify-between text-sm font-medium text-slate-700">
                    <span>History</span>
                    <span className="font-bold text-slate-900">{historyWeeks} weeks</span>
                  </label>
                  <input
                    type="range"
                    min={1}
                    max={12}
                    step={1}
                    value={historyWeeks}
                    onChange={(event) => setHistoryWeeks(Number(event.target.value))}
                    className="w-full accent-slate-500"
                  />
                </div>
                <div>
                  <label className="mb-2 flex items-center justify-between text-sm font-medium text-slate-700">
                    <span>Hypothetical staff change</span>
                    <span
                      className={`font-bold ${whatIfStaffDelta > 0 ? 'text-[#a7c0be]' : whatIfStaffDelta < 0 ? 'text-[#a84b2a]' : 'text-slate-900'}`}
                    >
                      {whatIfStaffDelta > 0 ? `+${whatIfStaffDelta}` : whatIfStaffDelta}
                    </span>
                  </label>
                  <input
                    type="range"
                    min={-5}
                    max={15}
                    step={1}
                    value={whatIfStaffDelta}
                    onChange={(event) => setWhatIfStaffDelta(Number(event.target.value))}
                    className="w-full accent-[#a7c0be]"
                  />
                </div>
                <div>
                  <label className="mb-2 flex items-center justify-between text-sm font-medium text-slate-700">
                    <span>Hypothetical client volume change</span>
                    <span
                      className={`font-bold ${whatIfClientDelta > 0 ? 'text-[#a7c0be]' : whatIfClientDelta < 0 ? 'text-[#a84b2a]' : 'text-slate-900'}`}
                    >
                      {whatIfClientDelta > 0 ? `+${whatIfClientDelta}` : whatIfClientDelta}
                    </span>
                  </label>
                  <input
                    type="range"
                    min={-10}
                    max={20}
                    step={1}
                    value={whatIfClientDelta}
                    onChange={(event) => setWhatIfClientDelta(Number(event.target.value))}
                    className="w-full accent-[#352b43]"
                  />
                </div>
              </div>
            </div>

            <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
              <h2 className="mb-4 text-xl font-bold text-slate-900">Staffing Capacity by Week</h2>
              <p className="mb-4 text-sm text-slate-600">
                New external hires spend their first {STAFF_TRAINING_PERIOD_DAYS} days in training. They appear as Staff in Training but are not counted as live projected staff until training is complete.
              </p>
              <div className="overflow-x-auto">
                <div className="min-w-[1100px]">
                  <ResponsiveContainer width="100%" height={350}>
                    <ComposedChart data={projectionChartData} margin={{ left: 0, right: 0, top: 5, bottom: 5 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                      <XAxis dataKey="week" stroke="#64748b" />
                      <YAxis stroke="#64748b" />
                      <Tooltip content={<ProjectionTooltip />} />
                      <Legend />
                      {historicalProjectionWeeks.length > 0 && (
                        <ReferenceArea
                          x1={historicalProjectionWeeks[0].week}
                          x2={currentProjectionWeek?.week || historicalProjectionWeeks[historicalProjectionWeeks.length - 1].week}
                          fill="#e2e8f0"
                          fillOpacity={0.7}
                          stroke="none"
                          ifOverflow="extendDomain"
                        />
                      )}
                  <Area
                    dataKey="requiredStaff"
                    stackId="staffAboveRequired"
                    stroke="none"
                    fill="transparent"
                    legendType="none"
                    isAnimationActive={false}
                  />
                  <Area
                    dataKey="staffAboveRequired"
                    stackId="staffAboveRequired"
                    stroke="none"
                    fill="#a7c0be"
                    fillOpacity={0.28}
                    legendType="none"
                    isAnimationActive={false}
                  />
                  <Area
                    dataKey="requiredStaff"
                    stackId="staffBelowRequired"
                    stroke="none"
                    fill="transparent"
                    legendType="none"
                    isAnimationActive={false}
                  />
                  <Area
                    dataKey="staffBelowRequired"
                    stackId="staffBelowRequired"
                    stroke="none"
                    fill="#a84b2a"
                    fillOpacity={0.28}
                    legendType="none"
                    isAnimationActive={false}
                  />
                      <Line type="monotone" dataKey="projectedStudents" stroke="#a84b2a" strokeWidth={3} name="Projected Students" />
                      <Line type="monotone" dataKey="requiredStaff" stroke="#a7c0be" strokeWidth={2} name="Required Staff" />
                      <Line type="monotone" dataKey="projectedStaff" stroke="#352b43" strokeWidth={3} name="Live Projected Staff" />
                      <Line type="monotone" dataKey="staffInTraining" stroke="#d9a441" strokeWidth={2} strokeDasharray="5 5" name="Staff in Training" />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-6 lg:grid-cols-4">
              <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <h3 className="mb-4 text-lg font-bold text-slate-900">Discharge Events</h3>
                <div className="space-y-2">
                  {(() => {
                    const events = filteredCurrentClients
                      .filter((client) => {
                        const dischargeDate = client['Discharge Date'];
                        if (!dischargeDate) return false;

                        const daysUntilDischarge = dayDifference(new Date(), dischargeDate);
                        return daysUntilDischarge !== null && daysUntilDischarge >= 0 && daysUntilDischarge <= projectionWeeks * 7;
                      })
                      .sort((left, right) => String(left.Title || '').localeCompare(String(right.Title || ''), undefined, { sensitivity: 'base' }));

                    if (events.length === 0) {
                      return (
                        <p className="text-sm text-slate-500">
                          No discharges in the next {projectionWeeks * 7} days ({filteredCurrentClients.length} current
                          clients checked).
                        </p>
                      );
                    }

                    return events.map((client, index) => {
                      const daysUntilDischarge = dayDifference(new Date(), client['Discharge Date']);
                      return (
                        <div key={client.id ?? index} className="rounded-lg border-2 border-[#a84b2a] bg-[#b8b5b3]/35 p-3">
                          <p className="font-medium text-slate-900">{client.Title}</p>
                          <p className="text-sm text-slate-600">
                            {client.Services} • {client['Staffing Ratio']}
                          </p>
                          <p className="mt-1 text-xs text-[#a84b2a]">Discharges in {daysUntilDischarge} days</p>
                        </div>
                      );
                    });
                  })()}
                </div>
              </div>

              <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <h3 className="mb-4 text-lg font-bold text-slate-900">Enrollment Events</h3>
                <div className="space-y-2">
                  {filteredIntakeData
                    .map((item) => ({
                      item,
                      daysUntilStart: dayDifference(new Date(), item['Start Date'] || item['Tentative Start Date'])
                    }))
                    .filter(
                      ({ item, daysUntilStart }) =>
                        (item['Tentative Start Date'] || item['Start Date']) &&
                        daysUntilStart !== null &&
                        daysUntilStart >= 0 &&
                        daysUntilStart <= projectionWeeks * 7
                    )
                    .sort((left, right) => String(left.item.Title || '').localeCompare(String(right.item.Title || ''), undefined, { sensitivity: 'base' }))
                    .map(({ item, daysUntilStart }, index) => (
                      <div key={item.id ?? index} className="rounded-lg border-2 border-[#a84b2a] bg-[#a7c0be]/35 p-3">
                        <p className="font-medium text-slate-900">{item.Title}</p>
                        <p className="text-sm text-slate-600">
                          {item.Services} • {item['Staffing Ratio']}
                        </p>
                        <p className="mt-1 text-xs text-[#a84b2a]">Starts in {daysUntilStart} days</p>
                      </div>
                    ))}
                </div>
              </div>

              <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <h3 className="mb-4 text-lg font-bold text-slate-900">Upcoming Staff</h3>
                <div className="space-y-2">
                  {prospectiveStaff.length === 0 ? (
                    <p className="text-sm text-slate-500">No upcoming staff found.</p>
                  ) : (
                    prospectiveStaff
                      .filter((person) => person.startDate)
                      .map((person) => ({
                        person,
                        daysUntilStart: Math.floor((new Date(person.startDate) - new Date()) / (1000 * 60 * 60 * 24))
                      }))
                      .filter(({ daysUntilStart }) => daysUntilStart <= projectionWeeks * 7)
                      .sort((left, right) => String(left.person.name || '').localeCompare(String(right.person.name || ''), undefined, { sensitivity: 'base' }))
                      .map(({ person, daysUntilStart }, index) => (
                        <div key={person.id ?? index} className="rounded-lg border-2 border-[#352b43] bg-[#beb7c8]/30 p-3">
                          <p className="font-medium text-slate-900">
                            {person.name}
                          </p>
                          <p className="text-sm text-slate-600">{person.position}{person.status ? ` • ${person.status}` : ''}</p>
                          <p className="mt-1 text-xs text-[#352b43]">
                            {daysUntilStart > 0 ? `Starts in ${daysUntilStart} days` : 'Started'}
                          </p>
                        </div>
                      ))
                  )}
                </div>
              </div>

              <div className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm">
                <h3 className="mb-4 text-lg font-bold text-slate-900">Upcoming Staff End Dates</h3>
                <div className="space-y-2">
                  {upcomingStaffEndDates.length === 0 ? (
                    <p className="text-sm text-slate-500">No upcoming staff end dates found.</p>
                  ) : (
                    alphabetizeNames(upcomingStaffEndDates).map((person, index) => (
                      <div key={person.id ?? index} className="rounded-lg border-2 border-[#d9a441] bg-[#fff4d6] p-3">
                        <p className="font-medium text-slate-900">
                          {person.name}{person.newPosition ? ` (Transfer: ${person.newPosition})` : ''}
                        </p>
                        <p className="text-sm text-slate-600">{person.role || 'Staff'} • {person.isActive ? 'Active' : 'Inactive'}</p>
                        <p className="mt-1 text-xs text-[#9a6b08]">
                          Ends in {person.daysUntilEnd} days ({formatDate(person.endDate)})
                        </p>
                      </div>
                    ))
                  )}
                </div>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function TimelineTooltip({ active, payload }) {
  if (!active || !payload?.length) return null;
  const row = payload.find((entry) => entry.dataKey === 'spanMs')?.payload;
  if (!row) return null;

  return (
    <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 shadow-sm">
      <p className="text-sm font-medium text-slate-900">{row.name}</p>
      <p className="text-xs text-slate-600">Inquiry: {row.inquiryDateLabel}</p>
      {row.isOpenEnded ? (
        <p className="text-xs text-slate-600">No start date set yet — bar runs to today</p>
      ) : (
        <p className="text-xs text-slate-600">
          {row.isTentative ? 'Tentative start' : 'Start'}: {row.startDateLabel}
        </p>
      )}
      <p className="mt-1 text-xs font-semibold text-[#3f425e]">
        {row.isOpenEnded ? `${row.daysToStart} days from inquiry to today` : `${row.daysToStart} days from inquiry to start`}
      </p>
    </div>
  );
}

function ProjectionTooltip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  const row = payload[0].payload;
  const studentsLabel = row.isHistorical ? 'Actual Client Numbers' : 'Projected Client Numbers';
  const staffLabel = row.isHistorical ? 'Actual Staff Numbers' : 'Projected Staff Numbers';

  return (
    <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 shadow-sm">
      <p className="text-sm font-medium text-slate-900">{label}</p>
      <p className="text-xs text-slate-600">
        {studentsLabel}: <span className="font-semibold text-slate-900">{row.projectedStudents}</span>
      </p>
      <p className="text-xs text-slate-600">
        Required Staff: <span className="font-semibold text-slate-900">{row.requiredStaff}</span>
      </p>
      <p className="text-xs text-slate-600">
        {staffLabel}: <span className="font-semibold text-slate-900">{row.projectedStaff}</span>
      </p>
      <p className="text-xs text-slate-600">
        Staff in Training: <span className="font-semibold text-slate-900">{row.staffInTraining}</span>
      </p>
      <p className={`mt-1 text-xs font-semibold ${row.staffingGap < 0 ? 'text-[#a84b2a]' : 'text-slate-900'}`}>
        Staffing Gap: {row.staffingGap > 0 ? `+${row.staffingGap}` : row.staffingGap}
      </p>
    </div>
  );
}

function StatCard({ title, value, subtitle, color }) {
  const colorMap = {
    blue: 'border-[#b8b5b3] bg-[#f1f1f1] text-[#3f425e]',
    green: 'border-[#a7c0be] bg-[#edf3f2] text-[#3f425e]',
    purple: 'border-[#beb7c8] bg-[#beb7c8]/30 text-[#352b43]',
    orange: 'border-[#d8a695] bg-[#f8e9e3] text-[#a84b2a]'
  };

  return (
    <div className={`rounded-lg border p-6 shadow-sm ${colorMap[color]}`}>
      <p className="text-sm font-medium opacity-75">{title}</p>
      <p className="mt-2 text-4xl font-bold">{value}</p>
      <p className="mt-2 text-xs opacity-75">{subtitle}</p>
    </div>
  );
}

function FilterSelect({ label, options, value, onChange }) {
  return (
    <div>
      <label className="mb-2 block text-sm font-medium text-slate-700">{label}</label>
      <select
        value={value || ''}
        onChange={(event) => onChange(event.target.value || null)}
        className="w-full rounded-lg border border-[#b8b5b3] bg-white px-3 py-2 text-slate-900 hover:border-[#3f425e] focus:outline-none focus:ring-2 focus:ring-[#352b43]"
      >
        <option value="">All {label}s</option>
        {options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    </div>
  );
}

function StatusBadge({ status }) {
  const colorMap = {
    'Initial Inquiry': 'bg-[#beb7c8]/30 text-[#352b43]',
    'Contacted Referral': 'bg-[#f1f1f1] text-[#3f425e]',
    'Sent Intake Email': 'bg-[#edf3f2] text-[#3f425e]',
    'Intake Completed': 'bg-[#e5e0eb] text-[#352b43]',
    'Awaiting Clinician Review': 'bg-[#fff4d6] text-[#9a6b08]',
    Enrolled: 'bg-[#f8e9e3] text-[#a84b2a]'
  };

  return (
    <span className={`rounded-full px-3 py-1 text-xs font-semibold ${colorMap[status] || 'bg-slate-100 text-slate-800'}`}>
      {status}
    </span>
  );
}

function DetailPanel({ item }) {
  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-4 md:grid-cols-4">
        <DetailField label="Inquiry Date" value={formatDate(item['Inquiry Date'])} />
        <DetailField label="Referral Source" value={item['Referral Source'] || 'N/A'} />
        <DetailField label="School District" value={item['School District'] || 'N/A'} />
        <DetailField label="Days to Start" value={item.DaysToStart >= 0 ? `${item.DaysToStart}d` : 'Overdue'} />
        <DetailField label="Parent/Guardian" value={item['Parent/Guardian Name'] || 'N/A'} />
        <DetailField label="Contact" value={item['Contact Num'] || 'N/A'} />
        <DetailField label="Email" value={item.Email || 'N/A'} />
        <DetailField label="DOB" value={formatDate(item['Client DOB']) || 'N/A'} />
        <DetailField label="BCBA" value={item.BCBA || 'N/A'} />
        <DetailField label="RMHS Eligible" value={item['RMHS Elig'] || 'N/A'} />
      </div>
      {item.Notes && (
        <div className="rounded-lg bg-slate-100 p-3">
          <p className="text-sm font-medium text-slate-700">Notes</p>
          <p className="mt-1 text-sm text-slate-600">{item.Notes}</p>
        </div>
      )}
    </div>
  );
}

function DetailField({ label, value }) {
  return (
    <div>
      <p className="text-xs font-medium uppercase tracking-wider text-slate-500">{label}</p>
      <p className="mt-1 text-sm font-medium text-slate-900">{value}</p>
    </div>
  );
}

function formatDate(dateString) {
  if (!dateString) return '';
  return new Date(dateString).toLocaleDateString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: '2-digit'
  });
}

function calculateAverageRatio(clients) {
  if (!clients.length) return 'N/A';
  const ratios = clients.map((client) => {
    const ratio = client['Staffing Ratio'];
    if (!ratio) return 1;
    const parts = ratio.split(':');
    return Number.parseInt(parts[1], 10) || 1;
  });
  const average = (ratios.reduce((sum, value) => sum + value, 0) / ratios.length).toFixed(1);
  return `1:${average}`;
}

