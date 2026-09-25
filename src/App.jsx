import React, { useState, useEffect, useMemo, useCallback } from 'react';
import { 
  INITIAL_MEMBERS, 
  INITIAL_MOVEMENTS, 
  INITIAL_CHI_BO, 
  INITIAL_BRANCH_DETAILS,
  INITIAL_EVALUATIONS, 
  INITIAL_FEES,
  INITIAL_MEETING_SCHEDULES
} from './data';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import {
  Users,
  Calendar,
  Award,
  FileSpreadsheet,
  Printer,
  Plus,
  Search,
  CheckCircle2,
  AlertCircle,
  Clock,
  ShieldCheck,
  Sun,
  Moon,
  Sparkles,
  Download,
  Send,
  CheckCheck,
  Share2,
  Copy,
  Check,
  MapPin,
  Archive,
  BookOpen,
  FileCheck2,
  FileText,
  RefreshCw,
  Settings,
  Database,
  ExternalLink,
  Wifi,
  WifiOff,
  Edit,
  Save,
  CalendarDays
} from 'lucide-react';

// Danh sách tên chuẩn 13 Chi bộ theo mẫu HD 01-HD/TU
const STANDARD_13_BRANCHES = [
  { id: 1, name: 'Chi bộ 1', shortName: 'Chi bộ 1', sl: 18, biThu: 'Đặng Thanh Phê', sdt: '0918.233.352' },
  { id: 2, name: 'Chi bộ 2', shortName: 'Chi bộ 2', sl: 9, biThu: 'Đào Minh Phúc', sdt: '0943.685.678' },
  { id: 3, name: 'Chi bộ 3', shortName: 'Chi bộ 3', sl: 9, biThu: 'Nguyễn Hữu Thọ', sdt: '0915.755.761' },
  { id: 4, name: 'Chi bộ 4', shortName: 'Chi bộ 4', sl: 8, biThu: 'Dương Thành Sự', sdt: '0919.999.182' },
  { id: 5, name: 'Chi bộ 5', shortName: 'Chi bộ 5', sl: 12, biThu: 'Trương Cáo', sdt: '0939.949.499' },
  { id: 6, name: 'Chi bộ 6', shortName: 'Chi bộ 6', sl: 10, biThu: 'Bùi Phước Lan', sdt: '0919.922.773' },
  { id: 7, name: 'Chi bộ 7', shortName: 'Chi bộ 7', sl: 8, biThu: 'Ngô Chí Trung', sdt: '0911.658.288' },
  { id: 8, name: 'Chi bộ 8', shortName: 'Chi bộ 8', sl: 12, biThu: 'Diệp Trọng Danh', sdt: '0913.125.981' },
  { id: 9, name: 'Chi bộ 9', shortName: 'Chi bộ 9', sl: 12, biThu: 'Trần Thị Thu Thanh Thủy', sdt: '0989.625.878' },
  { id: 10, name: 'Chi bộ 10', shortName: 'Chi bộ 10', sl: 9, biThu: 'Nguyễn Phúc Xuân Thụy', sdt: '0918.823.001' },
  { id: 11, name: 'Chi bộ 11', shortName: 'Chi bộ 11', sl: 12, biThu: 'Trần Đình Chinh', sdt: '0918.529.567' },
  { id: 12, name: 'Chi bộ 12', shortName: 'Chi bộ 12', sl: 15, biThu: 'Hà Vĩnh Tân', sdt: '0913.784.878' },
  { id: 13, name: 'Chi bộ KP', shortName: 'Chi bộ KP', sl: 23, biThu: 'Nguyễn Trung Tiến', sdt: '0918.666.888' }
];

// Hàm format ngày và tự động tính thứ trong tuần (Tiếng Việt)
const formatVietnameseDateWithDay = (dateString) => {
  if (!dateString) return '';
  const [yyyy, mm, dd] = dateString.split('-');
  if (!yyyy || !mm || !dd) return dateString;
  const dateObj = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
  const days = ['Chủ nhật', 'thứ Hai', 'thứ Ba', 'thứ Tư', 'thứ Năm', 'thứ Sáu', 'thứ Bảy'];
  const dayName = days[dateObj.getDay()];
  return `${dd}/${mm}/${yyyy} (${dayName})`;
};

const formatVietnameseDateOnly = (dateString) => {
  if (!dateString) return '';
  const [yyyy, mm, dd] = dateString.split('-');
  if (!yyyy || !mm || !dd) return dateString;
  return `${dd}/${mm}/${yyyy}`;
};

export default function App() {
  const [darkMode, setDarkMode] = useState(false);

  // Lấy tham số động từ URL (mode, month, year, chibo, api)
  const urlParams = new URLSearchParams(window.location.search);
  const isPublicRegisterMode = urlParams.get('mode') === 'register';
  const isPublicReportMode = urlParams.get('mode') === 'report';
  const paramChiBoId = urlParams.get('chibo');
  const paramMonth = urlParams.get('month');
  const paramYear = urlParams.get('year');
  const paramApi = urlParams.get('api');

  const DEFAULT_GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxUk8NpDGVkRpJOumVGubycnxmc7PVyItvNnv8qjlzF3sgeY5O7aiWPDh_klwPoULzJ/exec';

  // Google Sheets Apps Script Web App URL
  const [googleScriptUrl, setGoogleScriptUrl] = useState(() => {
    if (paramApi) {
      localStorage.setItem('qltt_google_script_url', paramApi);
      return paramApi;
    }
    return localStorage.getItem('qltt_google_script_url') || DEFAULT_GOOGLE_SCRIPT_URL;
  });

  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState('');
  const [isApiModalOpen, setIsApiModalOpen] = useState(false);
  const [apiInputUrl, setApiInputUrl] = useState(googleScriptUrl);
  const [saveToast, setSaveToast] = useState('');

  // Core Data States
  const [members, setMembers] = useState(() => {
    const saved = localStorage.getItem('qltt_party_members_v3');
    return saved ? JSON.parse(saved) : INITIAL_MEMBERS;
  });

  const [movements, setMovements] = useState(() => {
    const saved = localStorage.getItem('qltt_party_movements_v3');
    return saved ? JSON.parse(saved) : INITIAL_MOVEMENTS;
  });

  const [chiBoList] = useState(INITIAL_CHI_BO);
  const [branchDetails] = useState(INITIAL_BRANCH_DETAILS);

  // Lịch sinh hoạt 13 chi bộ theo tháng (Bảng Đăng Ký & Thông Báo Lịch Họp)
  const [meetingSchedules, setMeetingSchedules] = useState(() => {
    const saved = localStorage.getItem('qltt_party_meetings_v3');
    return saved ? JSON.parse(saved) : INITIAL_MEETING_SCHEDULES;
  });

  // Biểu Đánh Giá Xếp Loại Chất Lượng Sinh Hoạt Chi Bộ Hằng Tháng (ĐGXL) theo HD 01-HD/TU
  const [dgxlData, setDgxlData] = useState(() => {
    const saved = localStorage.getItem('qltt_party_dgxl_v3');
    return saved ? JSON.parse(saved) : {};
  });

  // Ghi chú biến động tổng thể chân trang ĐGXL
  const [dgxlOverallNotes, setDgxlOverallNotes] = useState(() => {
    const saved = localStorage.getItem('qltt_party_dgxl_notes');
    return saved ? JSON.parse(saved) : {};
  });

  // Biểu Tổng Hợp Số Liệu Báo Cáo 06 Tháng / Năm
  const [sixMonthsReports, setSixMonthsReports] = useState(() => {
    const saved = localStorage.getItem('qltt_party_sixmonths_v3');
    return saved ? JSON.parse(saved) : {};
  });

  // Lưu trữ Lịch sử Thông Báo Lịch Họp đã ban hành
  const [archivedNotices, setArchivedNotices] = useState(() => {
    const saved = localStorage.getItem('qltt_party_archived_notices');
    return saved ? JSON.parse(saved) : [];
  });

  const [activeTab, setActiveTab] = useState('meetings');

  // Đọc tháng và năm linh hoạt từ URL nếu có (ví dụ ?month=10&year=2026), mặc định Tháng 10 Năm 2026
  const [selectedYear, setSelectedYear] = useState(() => paramYear ? Number(paramYear) : 2026);
  const [selectedMonth, setSelectedMonth] = useState(() => paramMonth ? Number(paramMonth) : 10);

  const [searchTerm, setSearchTerm] = useState('');
  const [filterChiBo, setFilterChiBo] = useState('Tất cả');

  // Modal Share Link
  const [isShareModalOpen, setIsShareModalOpen] = useState(false);
  const [shareTab, setShareTab] = useState('register'); // 'register' | 'report'
  const [copiedLink, setCopiedLink] = useState(false);
  const [copiedMessage, setCopiedMessage] = useState(false);

  // Print Mode State ('schedule' | 'dgxl')
  const [printDocType, setPrintDocType] = useState('schedule');

  // Modal Sửa Chi Tiết Cho Quản Trị Viên (Admin Quick Edit Modal)
  const [adminEditModal, setAdminEditModal] = useState({
    isOpen: false,
    type: 'meeting', // 'meeting' | 'dgxl'
    chiBoId: 1,
    chiBoName: 'Chi bộ 1',
    dateValue: '',
    dateFormatted: '',
    location: '',
    attendance: '18/18',
    score: 100,
    classification: 'Tốt',
    note: '',
    status: 'Đã thống nhất'
  });

  // Form tự đăng ký lịch họp của Đội (mode=register)
  const initialSelectedBranchId = paramChiBoId ? Number(paramChiBoId) : 1;
  const initialBranchObj = branchDetails.find(b => b.id === initialSelectedBranchId) || branchDetails[0];

  const [teamSelectChiBo, setTeamSelectChiBo] = useState(initialSelectedBranchId);
  const [teamDate, setTeamDate] = useState('');
  const [teamDayOfWeek, setTeamDayOfWeek] = useState('thứ Hai');
  const [teamLocation, setTeamLocation] = useState(initialBranchObj.diaDiem);
  const [teamNote, setTeamNote] = useState('');
  const [teamSubmitted, setTeamSubmitted] = useState(false);
  const [isSubmittingOnline, setIsSubmittingOnline] = useState(false);

  // Form tự báo cáo kết quả ĐGXL của Đội sau khi họp (mode=report)
  const [reportBranchId, setReportBranchId] = useState(initialSelectedBranchId);
  const [reportAttendance, setReportAttendance] = useState(`${initialBranchObj.sl}/${initialBranchObj.sl}`);
  const [reportMeetingDate, setReportMeetingDate] = useState('');
  const [reportScore, setReportScore] = useState(100);
  const [reportClassification, setReportClassification] = useState('Tốt');
  const [reportDeductionNote, setReportDeductionNote] = useState('');
  const [reportSubmitted, setReportSubmitted] = useState(false);

  useEffect(() => {
    const br = branchDetails.find(b => b.id === teamSelectChiBo);
    if (br) {
      setTeamLocation(br.diaDiem);
    }
  }, [teamSelectChiBo, branchDetails]);

  useEffect(() => {
    const br = branchDetails.find(b => b.id === reportBranchId);
    if (br) {
      setReportAttendance(`${br.sl}/${br.sl}`);
    }
  }, [reportBranchId, branchDetails]);

  useEffect(() => {
    localStorage.setItem('qltt_party_members_v3', JSON.stringify(members));
  }, [members]);
  useEffect(() => {
    localStorage.setItem('qltt_party_movements_v3', JSON.stringify(movements));
  }, [movements]);
  useEffect(() => {
    localStorage.setItem('qltt_party_meetings_v3', JSON.stringify(meetingSchedules));
  }, [meetingSchedules]);
  useEffect(() => {
    localStorage.setItem('qltt_party_dgxl_v3', JSON.stringify(dgxlData));
  }, [dgxlData]);
  useEffect(() => {
    localStorage.setItem('qltt_party_dgxl_notes', JSON.stringify(dgxlOverallNotes));
  }, [dgxlOverallNotes]);
  useEffect(() => {
    localStorage.setItem('qltt_party_sixmonths_v3', JSON.stringify(sixMonthsReports));
  }, [sixMonthsReports]);
  useEffect(() => {
    localStorage.setItem('qltt_party_archived_notices', JSON.stringify(archivedNotices));
  }, [archivedNotices]);
  useEffect(() => {
    if (googleScriptUrl) {
      localStorage.setItem('qltt_google_script_url', googleScriptUrl);
    }
  }, [googleScriptUrl]);

  // HÀM ĐỒNG BỘ DỮ LIỆU TỪ GOOGLE SHEETS VỀ ỨNG DỤNG
  const fetchFromGoogleSheets = useCallback(async (customUrl) => {
    const targetUrl = customUrl || googleScriptUrl;
    if (!targetUrl) return;

    try {
      setIsSyncing(true);
      const res = await fetch(targetUrl);
      const data = await res.json();

      if (data && data.status === 'success') {
        // 1. Cập nhật Lịch họp
        if (data.meetings && data.meetings.length > 0) {
          setMeetingSchedules(prev => {
            const nextState = { ...prev };
            data.meetings.forEach(item => {
              const key = `${item.nam}-${item.thang}`;
              const currentList = nextState[key] || branchDetails.map(b => ({
                chiBoId: b.id,
                chiBo: b.name,
                sl: b.sl,
                thoiGian: 'Chưa đăng ký',
                diaDiem: b.diaDiem,
                biThu: b.biThu,
                sdt: b.sdt,
                trangThai: 'Chờ đăng ký',
                ghiChu: ''
              }));

              nextState[key] = currentList.map(bItem => {
                if (bItem.chiBoId === Number(item.chiBoId)) {
                  return {
                    ...bItem,
                    thoiGian: item.thoiGian || bItem.thoiGian,
                    diaDiem: item.diaDiem || bItem.diaDiem,
                    ghiChu: item.ghiChu || bItem.ghiChu,
                    trangThai: 'Đã đăng ký (Chờ duyệt)'
                  };
                }
                return bItem;
              });
            });
            return nextState;
          });
        }

        // 2. Cập nhật ĐGXL
        if (data.dgxl && data.dgxl.length > 0) {
          setDgxlData(prev => {
            const nextState = { ...prev };
            data.dgxl.forEach(item => {
              const key = `${item.nam}-${item.thang}`;
              const currentList = nextState[key] || STANDARD_13_BRANCHES.map(b => ({
                chiBoId: b.id,
                chiBo: b.shortName,
                sl: '',
                ngayHop: '',
                diemDG: '',
                mucXepLoai: '',
                ghiChuTruDiem: ''
              }));

              nextState[key] = currentList.map(bItem => {
                if (bItem.chiBoId === Number(item.chiBoId)) {
                  return {
                    ...bItem,
                    sl: item.sl || bItem.sl,
                    ngayHop: item.ngayHop || bItem.ngayHop,
                    diemDG: item.diemDG !== undefined && item.diemDG !== '' ? Number(item.diemDG) : bItem.diemDG,
                    mucXepLoai: item.mucXepLoai || bItem.mucXepLoai,
                    ghiChuTruDiem: item.ghiChuTruDiem || bItem.ghiChuTruDiem
                  };
                }
                return bItem;
              });
            });
            return nextState;
          });
        }

        setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
      }
    } catch (err) {
      console.error('Lỗi đồng bộ Google Sheets:', err);
    } finally {
      setIsSyncing(false);
    }
  }, [googleScriptUrl, branchDetails]);

  // Tự động đồng bộ khi mở trang nếu đã cấu hình Google Script URL
  useEffect(() => {
    if (googleScriptUrl) {
      fetchFromGoogleSheets(googleScriptUrl);
    }
  }, [googleScriptUrl, fetchFromGoogleSheets]);

  const currentKey = `${selectedYear}-${selectedMonth}`;

  // Ghi chú chân trang ĐGXL của tháng hiện tại
  const currentMonthNote = dgxlOverallNotes[currentKey] !== undefined 
    ? dgxlOverallNotes[currentKey] 
    : 'Tổng số: 157 đảng viên (trong đó, có 01 đv chuyển sinh hoạt tạm thời về Trường Chính trị của CB1)./.';

  const handleUpdateOverallNote = (val) => {
    setDgxlOverallNotes(prev => ({
      ...prev,
      [currentKey]: val
    }));
  };

  const currentMonthSchedules = useMemo(() => {
    if (meetingSchedules[currentKey]) {
      return meetingSchedules[currentKey].map(item => {
        const b = branchDetails.find(br => br.id === item.chiBoId) || {};
        return {
          ...item,
          biThu: item.biThu || b.biThu || '',
          sdt: item.sdt || b.sdt || '',
          diaDiem: item.diaDiem || b.diaDiem || ''
        };
      });
    }
    return branchDetails.map(b => ({
      chiBoId: b.id,
      chiBo: b.name,
      sl: b.sl,
      thoiGian: 'Chưa đăng ký',
      diaDiem: b.diaDiem,
      biThu: b.biThu,
      sdt: b.sdt,
      trangThai: 'Chờ đăng ký',
      ghiChu: ''
    }));
  }, [meetingSchedules, currentKey, branchDetails]);

  // HÀM LƯU 1 LỊCH HỌP LÊN GOOGLE SHEETS
  const saveMeetingToGoogleSheets = async (item, customMonth, customYear) => {
    if (!googleScriptUrl) return;
    const targetYear = customYear || selectedYear;
    const targetMonth = customMonth || selectedMonth;
    try {
      await fetch(googleScriptUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          action: 'register_meeting',
          nam: targetYear,
          thang: targetMonth,
          chiBoId: item.chiBoId,
          chiBo: item.chiBo,
          thoiGian: item.thoiGian,
          diaDiem: item.diaDiem,
          ghiChu: item.ghiChu || ''
        })
      });
    } catch (err) {
      console.error('Lỗi khi lưu lịch họp lên Google Sheets:', err);
    }
  };

  // HÀM LƯU TOÀN BỘ 13 CHI BỘ LỊCH HỌP LÊN GOOGLE SHEETS
  const handleSaveAllMeetingsToGoogleSheets = async () => {
    if (!googleScriptUrl) {
      setIsApiModalOpen(true);
      return;
    }
    setIsSyncing(true);
    try {
      for (const item of currentMonthSchedules) {
        await fetch(googleScriptUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'register_meeting',
            nam: selectedYear,
            thang: selectedMonth,
            chiBoId: item.chiBoId,
            chiBo: item.chiBo,
            thoiGian: item.thoiGian,
            diaDiem: item.diaDiem,
            ghiChu: item.ghiChu || ''
          })
        });
      }
      setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
      setSaveToast(`Đã lưu toàn bộ Lịch họp Tháng ${selectedMonth}/${selectedYear} lên Google Sheets!`);
      setTimeout(() => setSaveToast(''), 4000);
    } catch (err) {
      alert('Lỗi khi lưu lên Google Sheets: ' + err.message);
    } finally {
      setIsSyncing(false);
    }
  };

  // Quản trị viên cập nhật trực tiếp tại dòng bảng Lịch Họp
  const handleUpdateSchedule = (chiBoId, field, value) => {
    const updatedList = currentMonthSchedules.map(item => {
      if (item.chiBoId === chiBoId) {
        const nextItem = { ...item, [field]: value };
        if (field === 'thoiGian' && value && value !== 'Chưa đăng ký') {
          if (nextItem.trangThai === 'Chờ đăng ký') nextItem.trangThai = 'Đã đăng ký (Chờ duyệt)';
        }
        return nextItem;
      }
      return item;
    });
    setMeetingSchedules(prev => ({
      ...prev,
      [currentKey]: updatedList
    }));

    // Tự động lưu lên Google Sheets
    const modifiedItem = updatedList.find(i => i.chiBoId === chiBoId);
    if (modifiedItem) {
      saveMeetingToGoogleSheets(modifiedItem);
    }
  };

  // Quản trị viên chọn ngày từ DatePicker cho 1 Chi bộ tại Tab Lịch họp
  const handlePickDateForMeeting = (chiBoId, dateYMD) => {
    if (!dateYMD) return;
    const formatted = formatVietnameseDateWithDay(dateYMD);
    handleUpdateSchedule(chiBoId, 'thoiGian', formatted);
  };

  const handleApproveAndArchive = async () => {
    const updatedList = currentMonthSchedules.map(item => ({
      ...item,
      trangThai: 'Đã thống nhất'
    }));
    setMeetingSchedules(prev => ({
      ...prev,
      [currentKey]: updatedList
    }));

    const newArchiveItem = {
      id: 'TB_' + currentKey + '_' + Date.now(),
      nam: selectedYear,
      thang: selectedMonth,
      ngayBanHanh: new Date().toLocaleDateString('vi-VN'),
      nguoiKy: 'Bí thư Nguyễn Trung Tiến',
      tongSoDV: updatedList.reduce((sum, i) => sum + (Number(i.sl) || 0), 0),
      chiTiet: updatedList
    };

    setArchivedNotices(prev => {
      const filtered = prev.filter(a => !(a.nam === selectedYear && a.thang === selectedMonth));
      return [newArchiveItem, ...filtered];
    });

    // Đồng thời đồng bộ toàn bộ lên Google Sheets
    if (googleScriptUrl) {
      for (const item of updatedList) {
        saveMeetingToGoogleSheets(item);
      }
    }

    alert(`Đã duyệt thống nhất với Lãnh đạo và tự động lưu vào 'Kho Lưu Trữ Thông Báo'! Đồng chí có thể bấm 'Xuất File Word (.DOC)' hoặc 'In Trình Ký'.`);
  };

  // Dữ liệu ĐGXL chuẩn 13 chi bộ (Chi bộ 1 -> Chi bộ 12, Chi bộ KP)
  const currentMonthDGXL = useMemo(() => {
    if (dgxlData[currentKey]) {
      return dgxlData[currentKey];
    }
    // Khởi tạo theo danh sách chuẩn 13 chi bộ
    return STANDARD_13_BRANCHES.map(b => {
      const schedule = currentMonthSchedules.find(s => s.chiBoId === b.id) || {};
      let meetingDate = schedule.thoiGian || '';
      if (meetingDate && meetingDate.includes('(')) {
        meetingDate = meetingDate.split('(')[0].trim();
      }
      
      return {
        chiBoId: b.id,
        chiBo: b.shortName,
        sl: '', // Để trống chờ Đội nhập hoặc cán bộ nhập
        ngayHop: meetingDate || '',
        diemDG: '',
        mucXepLoai: '',
        ghiChuTruDiem: schedule.ghiChu || ''
      };
    });
  }, [dgxlData, currentKey, currentMonthSchedules]);

  // HÀM LƯU 1 DÒNG ĐGXL LÊN GOOGLE SHEETS
  const saveDGXLToGoogleSheets = async (item, customMonth, customYear) => {
    if (!googleScriptUrl) return;
    const targetYear = customYear || selectedYear;
    const targetMonth = customMonth || selectedMonth;
    try {
      await fetch(googleScriptUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain' },
        body: JSON.stringify({
          action: 'report_dgxl',
          nam: targetYear,
          thang: targetMonth,
          chiBoId: item.chiBoId,
          chiBo: item.chiBo,
          sl: item.sl || '',
          ngayHop: item.ngayHop || '',
          diemDG: item.diemDG !== '' ? Number(item.diemDG) : '',
          mucXepLoai: item.mucXepLoai || '',
          ghiChuTruDiem: item.ghiChuTruDiem || ''
        })
      });
    } catch (err) {
      console.error('Lỗi khi lưu ĐGXL lên Google Sheets:', err);
    }
  };

  // HÀM LƯU TOÀN BỘ 13 CHI BỘ ĐGXL LÊN GOOGLE SHEETS
  const handleSaveAllDGXLToGoogleSheets = async () => {
    if (!googleScriptUrl) {
      setIsApiModalOpen(true);
      return;
    }
    setIsSyncing(true);
    try {
      for (const item of currentMonthDGXL) {
        await fetch(googleScriptUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'report_dgxl',
            nam: selectedYear,
            thang: selectedMonth,
            chiBoId: item.chiBoId,
            chiBo: item.chiBo,
            sl: item.sl || '',
            ngayHop: item.ngayHop || '',
            diemDG: item.diemDG !== '' ? Number(item.diemDG) : '',
            mucXepLoai: item.mucXepLoai || '',
            ghiChuTruDiem: item.ghiChuTruDiem || ''
          })
        });
      }
      setLastSyncTime(new Date().toLocaleTimeString('vi-VN'));
      setSaveToast(`Đã lưu toàn bộ Bảng ĐGXL Tháng ${selectedMonth}/${selectedYear} lên Google Sheets!`);
      setTimeout(() => setSaveToast(''), 4000);
    } catch (err) {
      alert('Lỗi khi lưu ĐGXL lên Google Sheets: ' + err.message);
    } finally {
      setIsSyncing(false);
    }
  };

  // Quản trị viên cập nhật ĐGXL tại dòng bảng
  const handleUpdateDGXL = (chiBoId, field, value) => {
    const updated = currentMonthDGXL.map(item => {
      if (item.chiBoId === chiBoId) {
        const nextItem = { ...item, [field]: value };
        // Tự động xếp loại nếu người dùng nhập điểm
        if (field === 'diemDG') {
          if (value === '' || value === null || isNaN(value)) {
            nextItem.mucXepLoai = '';
          } else {
            const score = Number(value);
            if (score >= 90) nextItem.mucXepLoai = 'Tốt';
            else if (score >= 70) nextItem.mucXepLoai = 'Khá';
            else if (score >= 50) nextItem.mucXepLoai = 'Trung bình';
            else nextItem.mucXepLoai = 'Kém';
          }
        }
        return nextItem;
      }
      return item;
    });
    setDgxlData(prev => ({
      ...prev,
      [currentKey]: updated
    }));

    // Tự động lưu lên Google Sheets
    const modifiedItem = updated.find(i => i.chiBoId === chiBoId);
    if (modifiedItem) {
      saveDGXLToGoogleSheets(modifiedItem);
    }
  };

  // Quản trị viên chọn ngày từ DatePicker cho 1 Chi bộ tại Tab ĐGXL
  const handlePickDateForDGXL = (chiBoId, dateYMD) => {
    if (!dateYMD) return;
    const formatted = formatVietnameseDateOnly(dateYMD);
    handleUpdateDGXL(chiBoId, 'ngayHop', formatted);
  };

  // Mở Modal Chỉnh Sửa Chi Tiết Cho Quản Trị Viên (Admin Edit Modal)
  const handleOpenAdminEditModal = (type, chiBoId) => {
    const br = branchDetails.find(b => b.id === chiBoId) || branchDetails[0];
    const schedule = currentMonthSchedules.find(s => s.chiBoId === chiBoId) || {};
    const dgxl = currentMonthDGXL.find(d => d.chiBoId === chiBoId) || {};

    let initialDateFormatted = type === 'meeting' ? (schedule.thoiGian || '') : (dgxl.ngayHop || '');
    let initialDateYMD = '';
    
    // Thử trích xuất ngày YYYY-MM-DD từ chuỗi DD/MM/YYYY
    if (initialDateFormatted) {
      const match = initialDateFormatted.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (match) {
        const dd = match[1].padStart(2, '0');
        const mm = match[2].padStart(2, '0');
        const yyyy = match[3];
        initialDateYMD = `${yyyy}-${mm}-${dd}`;
      }
    }

    setAdminEditModal({
      isOpen: true,
      type: type,
      chiBoId: chiBoId,
      chiBoName: br.name,
      dateValue: initialDateYMD,
      dateFormatted: initialDateFormatted,
      location: schedule.diaDiem || br.diaDiem,
      attendance: dgxl.sl || `${br.sl}/${br.sl}`,
      score: dgxl.diemDG !== '' ? Number(dgxl.diemDG) : 100,
      classification: dgxl.mucXepLoai || 'Tốt',
      note: type === 'meeting' ? (schedule.ghiChu || '') : (dgxl.ghiChuTruDiem || ''),
      status: schedule.trangThai || 'Đã thống nhất'
    });
  };

  // Lưu từ Modal Chỉnh Sửa Quản Trị Viên
  const handleSaveAdminEditModal = async (e) => {
    e.preventDefault();
    const { type, chiBoId, dateFormatted, location, attendance, score, classification, note, status } = adminEditModal;

    if (type === 'meeting') {
      const updatedList = currentMonthSchedules.map(item => {
        if (item.chiBoId === chiBoId) {
          return {
            ...item,
            thoiGian: dateFormatted || item.thoiGian,
            diaDiem: location || item.diaDiem,
            ghiChu: note || '',
            trangThai: status || item.trangThai
          };
        }
        return item;
      });

      setMeetingSchedules(prev => ({
        ...prev,
        [currentKey]: updatedList
      }));

      const itemToSave = updatedList.find(i => i.chiBoId === chiBoId);
      if (itemToSave) {
        await saveMeetingToGoogleSheets(itemToSave);
      }
    } else {
      const updatedList = currentMonthDGXL.map(item => {
        if (item.chiBoId === chiBoId) {
          return {
            ...item,
            sl: attendance,
            ngayHop: dateFormatted || item.ngayHop,
            diemDG: score !== '' ? Number(score) : '',
            mucXepLoai: classification,
            ghiChuTruDiem: note || ''
          };
        }
        return item;
      });

      setDgxlData(prev => ({
        ...prev,
        [currentKey]: updatedList
      }));

      const itemToSave = updatedList.find(i => i.chiBoId === chiBoId);
      if (itemToSave) {
        await saveDGXLToGoogleSheets(itemToSave);
      }
    }

    setAdminEditModal(prev => ({ ...prev, isOpen: false }));
    setSaveToast(`Đã cập nhật và lưu lên Google Sheets thành công cho ${adminEditModal.chiBoName}!`);
    setTimeout(() => setSaveToast(''), 4000);
  };

  const sixMonthsData = useMemo(() => {
    const sixKey = `${selectedYear}-6M`;
    if (sixMonthsReports[sixKey]) {
      return sixMonthsReports[sixKey];
    }
    return branchDetails.map(b => ({
      chiBoId: b.id,
      chiBo: b.name,
      dangSo: b.sl,
      soLgDuHop: `${b.sl}/${b.sl} (Đạt 100%)`,
      phatBieuYkien: '100% đảng viên phát biểu',
      danVanKheo: `Thực hiện mô hình Dân vận khéo năm ${selectedYear}`,
      hocTapChuyenDe: `Học tập Bác Hồ, Bác Tôn về chăm lo đời sống Nhân dân`,
      docBaiViet: 'Sinh hoạt các mẩu chuyện kể về Bác Hồ',
      khSinhHoatCD: 'Tổ chức sinh hoạt chuyên đề Quý I, Quý II đúng quy định',
      knd: '0',
      ktgs: `Thực hiện Kế hoạch kiểm tra, giám sát năm ${selectedYear}`,
      ghiChu: '',
      khac: ''
    }));
  }, [sixMonthsReports, selectedYear, branchDetails]);

  const handleUpdateSixMonths = (chiBoId, field, value) => {
    const sixKey = `${selectedYear}-6M`;
    const updated = sixMonthsData.map(item => {
      if (item.chiBoId === chiBoId) {
        return { ...item, [field]: value };
      }
      return item;
    });
    setSixMonthsReports(prev => ({
      ...prev,
      [sixKey]: updated
    }));
  };

  // Submit từ form đăng ký lịch họp (mode=register) - Lưu Local + Gửi Google Sheets Online
  const handleTeamSubmitRegistration = async (e) => {
    e.preventDefault();
    if (!teamDate) {
      alert('Vui lòng chọn ngày họp!');
      return;
    }
    const [yyyy, mm, dd] = teamDate.split('-');
    const targetYear = Number(yyyy);
    const targetMonth = Number(mm);
    const targetKey = `${targetYear}-${targetMonth}`;
    const timeFormatted = `${dd}/${mm}/${yyyy} (${teamDayOfWeek})`;
    
    const targetBranch = branchDetails.find(b => b.id === teamSelectChiBo);
    const branchName = targetBranch ? targetBranch.name : `Chi bộ ${teamSelectChiBo}`;
    const loc = teamLocation || (targetBranch ? targetBranch.diaDiem : 'Tại đơn vị');

    // 1. Cập nhật state local
    setMeetingSchedules(prev => {
      const currentList = prev[targetKey] || branchDetails.map(b => ({
        chiBoId: b.id,
        chiBo: b.name,
        sl: b.sl,
        thoiGian: 'Chưa đăng ký',
        diaDiem: b.diaDiem,
        biThu: b.biThu,
        sdt: b.sdt,
        trangThai: 'Chờ đăng ký',
        ghiChu: ''
      }));

      const updatedList = currentList.map(item => {
        if (item.chiBoId === teamSelectChiBo) {
          return {
            ...item,
            thoiGian: timeFormatted,
            diaDiem: loc,
            ghiChu: teamNote || item.ghiChu || '',
            trangThai: 'Đã đăng ký (Chờ duyệt)'
          };
        }
        return item;
      });

      return {
        ...prev,
        [targetKey]: updatedList
      };
    });

    // 2. Gửi dữ liệu lên Google Sheets Online (nếu có URL)
    if (googleScriptUrl) {
      setIsSubmittingOnline(true);
      try {
        await fetch(googleScriptUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'register_meeting',
            nam: targetYear,
            thang: targetMonth,
            chiBoId: teamSelectChiBo,
            chiBo: branchName,
            thoiGian: timeFormatted,
            diaDiem: loc,
            ghiChu: teamNote || ''
          })
        });
      } catch (err) {
        console.error('Lỗi khi gửi lên Google Sheets:', err);
      } finally {
        setIsSubmittingOnline(false);
      }
    }

    setTeamSubmitted(true);
    setTimeout(() => setTeamSubmitted(false), 5000);
  };

  // Submit từ form báo cáo ĐGXL sau họp (mode=report) - Lưu Local + Gửi Google Sheets Online
  const handleTeamSubmitReport = async (e) => {
    e.preventDefault();
    let targetYear = selectedYear;
    let targetMonth = selectedMonth;
    let dateStr = reportMeetingDate;

    if (reportMeetingDate && reportMeetingDate.includes('-')) {
      const [yyyy, mm, dd] = reportMeetingDate.split('-');
      targetYear = Number(yyyy);
      targetMonth = Number(mm);
      dateStr = `${Number(dd)}/${Number(mm)}/${yyyy}`;
    }
    const targetKey = `${targetYear}-${targetMonth}`;
    const targetBranch = STANDARD_13_BRANCHES.find(b => b.id === reportBranchId);
    const branchName = targetBranch ? targetBranch.shortName : `Chi bộ ${reportBranchId}`;

    // 1. Cập nhật state local
    setDgxlData(prev => {
      const currentList = prev[targetKey] || STANDARD_13_BRANCHES.map(b => ({
        chiBoId: b.id,
        chiBo: b.shortName,
        sl: '',
        ngayHop: '',
        diemDG: '',
        mucXepLoai: '',
        ghiChuTruDiem: ''
      }));

      const updatedList = currentList.map(item => {
        if (item.chiBoId === reportBranchId) {
          return {
            ...item,
            sl: reportAttendance,
            ngayHop: dateStr || item.ngayHop,
            diemDG: Number(reportScore),
            mucXepLoai: reportClassification,
            ghiChuTruDiem: reportDeductionNote || item.ghiChuTruDiem || ''
          };
        }
        return item;
      });

      return {
        ...prev,
        [targetKey]: updatedList
      };
    });

    // 2. Gửi dữ liệu lên Google Sheets Online (nếu có URL)
    if (googleScriptUrl) {
      setIsSubmittingOnline(true);
      try {
        await fetch(googleScriptUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'text/plain' },
          body: JSON.stringify({
            action: 'report_dgxl',
            nam: targetYear,
            thang: targetMonth,
            chiBoId: reportBranchId,
            chiBo: branchName,
            sl: reportAttendance,
            ngayHop: dateStr || '',
            diemDG: Number(reportScore),
            mucXepLoai: reportClassification,
            ghiChuTruDiem: reportDeductionNote || ''
          })
        });
      } catch (err) {
        console.error('Lỗi khi gửi báo cáo ĐGXL lên Google Sheets:', err);
      } finally {
        setIsSubmittingOnline(false);
      }
    }

    setReportSubmitted(true);
    setTimeout(() => setReportSubmitted(false), 5000);
  };

  // XUẤT EXCEL ĐGXL CHUẨN 10 CỘT HƯỚNG DẪN 01-HD/TU
  const handleExportDGXLToExcel = () => {
    const dataRows = currentMonthDGXL.map((d, idx) => ({
      'STT (1)': idx + 1,
      'Tên đơn vị (2)': d.chiBo,
      'Số lượng đảng viên (3)': d.sl || '',
      'Ngày họp chi bộ (4)': d.ngayHop || '',
      'KQ số điểm được đánh giá (5)': d.diemDG || '',
      'Tốt (6)': d.mucXepLoai === 'Tốt' ? 'x' : '',
      'Khá (7)': d.mucXepLoai === 'Khá' ? 'x' : '',
      'Trung bình (8)': d.mucXepLoai === 'Trung bình' ? 'x' : '',
      'Kém (9)': d.mucXepLoai === 'Kém' ? 'x' : '',
      'Ghi chú (thuyết minh điểm trừ) (10)': d.ghiChuTruDiem || ''
    }));

    const ws = XLSX.utils.json_to_sheet(dataRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `BC_T${selectedMonth}_HD01`);
    XLSX.writeFile(wb, `${selectedMonth}. BC_T${selectedMonth}_xếp loại chi bộ theo HD 01 TU.xlsx`);
  };

  // XUẤT FILE WORD .DOC ĐGXL Y HỆT HÌNH ẢNH MẪU 100% 1 TRANG A4
  const handleExportDGXLWordDoc = () => {
    const totTot = currentMonthDGXL.filter(d => d.mucXepLoai === 'Tốt').length;
    const totKha = currentMonthDGXL.filter(d => d.mucXepLoai === 'Khá').length;
    const totTB = currentMonthDGXL.filter(d => d.mucXepLoai === 'Trung bình').length;
    const totKem = currentMonthDGXL.filter(d => d.mucXepLoai === 'Kém').length;

    const tableRows = currentMonthDGXL.map((item, idx) => `
      <tr style="height: 20px;">
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${idx + 1}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 11pt;">${item.chiBo}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.sl || ''}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.ngayHop || ''}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.diemDG !== '' ? item.diemDG : ''}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.mucXepLoai === 'Tốt' ? 'x' : ''}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.mucXepLoai === 'Khá' ? 'x' : ''}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.mucXepLoai === 'Trung bình' ? 'x' : ''}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.mucXepLoai === 'Kém' ? 'x' : ''}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 10pt;">${item.ghiChuTruDiem || ''}</td>
      </tr>
    `).join('');

    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>DANH SÁCH TỔNG HỢP ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG SINH HOẠT CHI BỘ THÁNG ${selectedMonth} NĂM ${selectedYear}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          @page Section1 {
            size: 21.0cm 29.7cm;
            margin: 1.0cm 1.2cm 1.0cm 1.8cm;
            mso-header-margin: 0pt;
            mso-footer-margin: 0pt;
            mso-paper-source: 0;
          }
          div.Section1 { page: Section1; }
          body {
            font-family: 'Times New Roman', serif;
            font-size: 13pt;
            line-height: 100%;
            mso-line-height-rule: exactly;
            color: black;
            margin: 0;
            padding: 0;
          }
          p, div {
            margin-top: 0pt !important;
            margin-bottom: 0pt !important;
            line-height: 100% !important;
            mso-line-height-rule: exactly;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 0pt;
            margin-bottom: 0pt;
          }
          td, th {
            mso-line-height-rule: exactly;
            line-height: 100%;
          }
        </style>
      </head>
      <body>
        <div class="Section1">
          <!-- Header Table -->
          <table style="width: 100%; border: none; margin-bottom: 4pt;">
            <tr>
              <td style="width: 54%; text-align: center; vertical-align: top; border: none; font-size: 11pt; white-space: nowrap; line-height: 100%;">
                ĐẢNG BỘ SỞ CÔNG THƯƠNG TỈNH AN GIANG<br/>
                <b>ĐẢNG ỦY CHI CỤC QUẢN LÝ THỊ TRƯỜNG</b><br/>
                *
              </td>
              <td style="width: 46%; text-align: center; vertical-align: top; border: none; font-size: 12pt; white-space: nowrap; line-height: 100%;">
                <b>ĐẢNG CỘNG SẢN VIỆT NAM</b><br/><br/>
                <i>An Giang, ngày 10 tháng ${String(selectedMonth).padStart(2, '0')} năm ${selectedYear}</i>
              </td>
            </tr>
          </table>

          <!-- Title -->
          <div style="text-align: center; margin-top: 4pt; margin-bottom: 6pt;">
            <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">DANH SÁCH</p>
            <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">TỔNG HỢP ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG SINH HOẠT CHI BỘ</p>
            <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">Tháng ${selectedMonth} NĂM ${selectedYear}</p>
            <p style="font-size: 11pt; margin: 0; line-height: 100%;">-----</p>
          </div>

          <!-- 10 Columns Table theo HD 01-HD/TU -->
          <table style="width: 100%; border: 1px solid black; font-size: 10.5pt; margin-top: 2pt; margin-bottom: 4pt;">
            <tr style="text-align: center; font-weight: bold; font-size: 11pt; background-color: #f2f2f2;">
              <td rowspan="2" style="border: 1px solid black; padding: 2px; width: 4%;">S<br/>T<br/>T</td>
              <td rowspan="2" style="border: 1px solid black; padding: 2px 4px; width: 16%;">Tên đơn vị</td>
              <td rowspan="2" style="border: 1px solid black; padding: 2px; width: 9%;">Số<br/>lượng<br/>đảng viên</td>
              <td rowspan="2" style="border: 1px solid black; padding: 2px; width: 10%;">Ngày<br/>họp<br/>chi<br/>bộ</td>
              <td colspan="5" style="border: 1px solid black; padding: 2px;">Mức đánh giá xếp loại sau buổi<br/>sinh hoạt chi bộ hằng tháng</td>
              <td rowspan="2" style="border: 1px solid black; padding: 2px 4px; width: 23%;">Ghi chú<br/>(thuyết minh điểm trừ)</td>
            </tr>
            <tr style="text-align: center; font-weight: bold; font-size: 10.5pt; background-color: #f2f2f2;">
              <td style="border: 1px solid black; padding: 2px; width: 9%;">KQ số<br/>điểm<br/>được<br/>đánh giá</td>
              <td style="border: 1px solid black; padding: 2px; width: 7%;">Tốt</td>
              <td style="border: 1px solid black; padding: 2px; width: 7%;">Khá</td>
              <td style="border: 1px solid black; padding: 2px; width: 7%;">Trung<br/>bình</td>
              <td style="border: 1px solid black; padding: 2px; width: 7%;">Kém</td>
            </tr>
            <tr style="text-align: center; font-style: italic; font-size: 9.5pt; background-color: #fafafa;">
              <td style="border: 1px solid black; padding: 1px;">1</td>
              <td style="border: 1px solid black; padding: 1px;">2</td>
              <td style="border: 1px solid black; padding: 1px;">3</td>
              <td style="border: 1px solid black; padding: 1px;">4</td>
              <td style="border: 1px solid black; padding: 1px;">5</td>
              <td style="border: 1px solid black; padding: 1px;">6</td>
              <td style="border: 1px solid black; padding: 1px;">7</td>
              <td style="border: 1px solid black; padding: 1px;">8</td>
              <td style="border: 1px solid black; padding: 1px;">9</td>
              <td style="border: 1px solid black; padding: 1px;">10</td>
            </tr>
            ${tableRows}
            <tr style="font-weight: bold; text-align: center; height: 20px;">
              <td colspan="2" style="border: 1px solid black; padding: 2px 4px; text-align: center;">Tổng cộng</td>
              <td style="border: 1px solid black; padding: 2px;"></td>
              <td style="border: 1px solid black; padding: 2px;"></td>
              <td style="border: 1px solid black; padding: 2px;"></td>
              <td style="border: 1px solid black; padding: 2px;">${totTot > 0 ? totTot : ''}</td>
              <td style="border: 1px solid black; padding: 2px;">${totKha > 0 ? totKha : ''}</td>
              <td style="border: 1px solid black; padding: 2px;">${totTB > 0 ? totTB : ''}</td>
              <td style="border: 1px solid black; padding: 2px;">${totKem > 0 ? totKem : ''}</td>
              <td style="border: 1px solid black; padding: 2px;"></td>
            </tr>
          </table>

          <!-- Footer Note -->
          <div style="font-size: 11pt; margin-top: 4pt; margin-bottom: 8pt; text-align: justify; line-height: 100%;">
            <b>* Ghi chú:</b> ${currentMonthNote}
          </div>

          <!-- Signatures Table -->
          <table style="width: 100%; border: none; margin-top: 8pt;">
            <tr>
              <td style="width: 50%; text-align: center; vertical-align: top; border: none; font-size: 12pt; line-height: 110%;">
                <b>NGƯỜI LẬP BIỂU</b><br/><br/><br/><br/><br/>
                <b>Trần Đình Chinh</b>
              </td>
              <td style="width: 50%; text-align: center; vertical-align: top; border: none; font-size: 12pt; line-height: 110%;">
                <b>T/M ĐẢNG ỦY</b><br/>
                BÍ THƯ<br/><br/><br/><br/>
                <b>Nguyễn Trung Tiến</b>
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword;charset=utf-8' });
    saveAs(blob, `9. BC_T${selectedMonth}_xếp loại chi bộ theo HD 01 TU.doc`);
  };

  // XUẤT FILE WORD .DOC THÔNG BÁO LỊCH HỌP KHÍT 100% 1 TRANG A4
  const handleExportWordDoc = () => {
    const tableRows = currentMonthSchedules.map((item, idx) => `
      <tr style="height: 22px;">
        <td style="border: 1px solid black; padding: 2px 3px; text-align: center; font-size: 11pt;">${idx + 1}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 11pt;">${item.chiBo}</td>
        <td style="border: 1px solid black; padding: 2px 3px; text-align: center; font-size: 11pt;">${item.sl}</td>
        <td style="border: 1px solid black; padding: 2px 4px; text-align: center; font-size: 11pt;">${item.thoiGian || ''}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 11pt;">${item.diaDiem || ''}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 11pt;">${item.biThu || ''} - ${item.sdt || ''}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 11pt;">${item.ghiChu || ''}</td>
      </tr>
    `).join('');

    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>THÔNG BÁO LỊCH SINH HOẠT CHI BỘ THÁNG ${selectedMonth} NĂM ${selectedYear}</title>
        <!--[if gte mso 9]>
        <xml>
          <w:WordDocument>
            <w:View>Print</w:View>
            <w:Zoom>100</w:Zoom>
            <w:DoNotOptimizeForBrowser/>
          </w:WordDocument>
        </xml>
        <![endif]-->
        <style>
          @page Section1 {
            size: 21.0cm 29.7cm;
            margin: 1.0cm 1.2cm 1.0cm 1.8cm;
            mso-header-margin: 0pt;
            mso-footer-margin: 0pt;
            mso-paper-source: 0;
          }
          div.Section1 { page: Section1; }
          body {
            font-family: 'Times New Roman', serif;
            font-size: 13pt;
            line-height: 100%;
            mso-line-height-rule: exactly;
            color: black;
            margin: 0;
            padding: 0;
          }
          p, div {
            margin-top: 0pt !important;
            margin-bottom: 0pt !important;
            line-height: 100% !important;
            mso-line-height-rule: exactly;
          }
          table {
            border-collapse: collapse;
            width: 100%;
            margin-top: 0pt;
            margin-bottom: 0pt;
          }
          td, th {
            mso-line-height-rule: exactly;
            line-height: 100%;
          }
        </style>
      </head>
      <body>
        <div class="Section1">
          <!-- Header Table -->
          <table style="width: 100%; border: none; margin-bottom: 4pt;">
            <tr>
              <td style="width: 44%; text-align: center; vertical-align: top; border: none; font-size: 11pt; white-space: nowrap; line-height: 100%;">
                ĐẢNG BỘ CỤC QUẢN LÝ THỊ TRƯỜNG<br/>
                <b>ĐẢNG ỦY BỘ PHẬN</b><br/>
                *<br/>
                Số &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-TB/ĐUBP
              </td>
              <td style="width: 56%; text-align: center; vertical-align: top; border: none; font-size: 12pt; white-space: nowrap; line-height: 100%;">
                <b>ĐẢNG CỘNG SẢN VIỆT NAM</b><br/><br/>
                <i>An Giang, ngày &nbsp;&nbsp;&nbsp; tháng ${String(selectedMonth).padStart(2, '0')} năm ${selectedYear}</i>
              </td>
            </tr>
          </table>

          <!-- Title -->
          <div style="text-align: center; margin-top: 4pt; margin-bottom: 6pt;">
            <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">THÔNG BÁO</p>
            <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">Lịch sinh hoạt lệ các chi bộ trực thuộc Tháng ${selectedMonth} năm ${selectedYear}</p>
            <p style="font-size: 11pt; margin: 0; line-height: 100%;">-----</p>
          </div>

          <!-- Content paragraph -->
          <div style="text-align: justify; font-size: 13pt; margin-bottom: 4pt; line-height: 105%;">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Căn cứ Quy chế làm việc của Đảng ủy bộ phận; Đảng ủy bộ phận thông báo lịch sinh hoạt lệ các chi bộ trực thuộc trong Tháng ${selectedMonth} năm ${selectedYear} cụ thể như sau:
          </div>

          <!-- Schedule Table -->
          <table style="width: 100%; border: 1px solid black; font-size: 11pt; margin-top: 2pt; margin-bottom: 6pt;">
            <tr style="text-align: center; font-weight: bold; background-color: #f2f2f2; height: 26px;">
              <td style="border: 1px solid black; padding: 2px 3px; width: 4%;">STT</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 14%;">Chi bộ</td>
              <td style="border: 1px solid black; padding: 2px 3px; width: 7%;">Số lượng</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 22%;">Thời gian sinh hoạt</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 15%;">Địa điểm</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 22%;">Bí thư, điện thoại</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 16%;">Ghi chú</td>
            </tr>
            ${tableRows}
          </table>

          <!-- Footer paragraph -->
          <div style="text-align: justify; font-size: 13pt; margin-bottom: 8pt; line-height: 105%;">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Đề nghị các đồng chí Đảng ủy viên phụ trách các chi bộ sắp xếp thời gian cùng dự họp với các chi bộ theo thời gian và địa điểm nêu trên./.
          </div>

          <!-- Signatures Table -->
          <table style="width: 100%; border: none; margin-top: 4pt;">
            <tr>
              <td style="width: 50%; vertical-align: top; border: none; font-size: 10.5pt; line-height: 100%;">
                <b><i>Nơi nhận:</i></b><br/>
                - Các ĐUV (dự);<br/>
                - 13 chi bộ trực thuộc (t/h);<br/>
                - Lưu: ĐUBP.
              </td>
              <td style="width: 50%; text-align: center; vertical-align: top; border: none; font-size: 12pt; line-height: 110%;">
                <b>T/M ĐẢNG ỦY BỘ PHẬN</b><br/>
                BÍ THƯ<br/><br/><br/><br/>
                <b>Nguyễn Trung Tiến</b>
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword;charset=utf-8' });
    saveAs(blob, `TB_Lich_Hop_Chi_Bo_Thang_${selectedMonth}_${selectedYear}.doc`);
  };

  const handlePrintSchedule = () => {
    setPrintDocType('schedule');
    setTimeout(() => {
      window.print();
    }, 150);
  };

  const handlePrintDGXL = () => {
    setPrintDocType('dgxl');
    setTimeout(() => {
      window.print();
    }, 150);
  };

  // URL Đăng ký & Báo cáo kèm API Web App
  const apiParamStr = googleScriptUrl ? `&api=${encodeURIComponent(googleScriptUrl)}` : '';
  const currentRegisterLink = `${window.location.origin}${window.location.pathname}?mode=register&month=${selectedMonth}&year=${selectedYear}${apiParamStr}`;
  const currentReportLink = `${window.location.origin}${window.location.pathname}?mode=report&month=${selectedMonth}&year=${selectedYear}${apiParamStr}`;

  const copyShareLink = (url) => {
    navigator.clipboard.writeText(url);
    setCopiedLink(true);
    setTimeout(() => setCopiedLink(false), 3000);
  };

  const copyZaloMessage = (type) => {
    let msg = '';
    if (type === 'register') {
      msg = `Kính gửi các đồng chí Bí thư Chi bộ trực thuộc Đảng bộ QLTT,\n\nĐề nghị các Chi bộ thực hiện đăng ký lịch họp sinh hoạt chi bộ Tháng ${selectedMonth}/${selectedYear} trước ngày 20 theo đường link dưới đây:\n👉 ${currentRegisterLink}\n\nTrân trọng cảm ơn!`;
    } else {
      msg = `Kính gửi các đồng chí Bí thư Chi bộ trực thuộc Đảng bộ QLTT,\n\nSau khi hoàn thành buổi sinh hoạt chi bộ Tháng ${selectedMonth}/${selectedYear}, đề nghị các Chi bộ gửi Báo cáo Đánh giá, xếp loại sinh hoạt chi bộ (theo Hướng dẫn 01-HD/TU) tại đường link dưới đây để Đảng ủy tổng hợp:\n👉 ${currentReportLink}\n\nTrân trọng cảm ơn!`;
    }
    navigator.clipboard.writeText(msg);
    setCopiedMessage(true);
    setTimeout(() => setCopiedMessage(false), 3000);
  };

  // GIAO DIỆN DÀNH CHO CÁC ĐỘI TỰ ĐĂNG KÝ LỊCH HỌP (mode=register)
  if (isPublicRegisterMode) {
    const selectedBranchInfo = branchDetails.find(b => b.id === teamSelectChiBo) || branchDetails[0];

    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="bg-white border border-slate-200 rounded-2xl w-full max-w-xl shadow-xl overflow-hidden">
          <div className="bg-red-700 p-6 text-white text-center relative">
            <div className="w-12 h-12 mx-auto rounded-xl bg-yellow-400 text-red-900 flex items-center justify-center font-black text-2xl shadow mb-2">
              ★
            </div>
            <h2 className="text-base font-bold uppercase tracking-wide">ĐẢNG ỦY CHI CỤC QUẢN LÝ THỊ TRƯỜNG</h2>
            <h3 className="text-lg font-black mt-1">PHIẾU ĐĂNG KÝ LỊCH SINH HOẠT CHI BỘ</h3>
            <p className="text-sm font-semibold text-yellow-300 mt-1">Thực hiện đăng ký trước ngày 20 • Tháng {selectedMonth} năm {selectedYear}</p>
          </div>

          <form onSubmit={handleTeamSubmitRegistration} className="p-6 space-y-4">
            {teamSubmitted ? (
              <div className="p-5 rounded-xl bg-emerald-50 border border-emerald-200 text-center space-y-2">
                <CheckCircle2 className="w-12 h-12 text-emerald-600 mx-auto" />
                <h4 className="font-bold text-emerald-800 text-base">Đăng ký thành công Tháng {selectedMonth}/{selectedYear}!</h4>
                <p className="text-xs text-emerald-700">
                  Lịch họp của <b>{selectedBranchInfo.name}</b> vào ngày <b>{teamDate} ({teamDayOfWeek})</b> tại <b>{teamLocation}</b> đã được lưu vào hệ thống máy chủ Đảng ủy.
                </p>
                <div className="pt-2">
                  <button
                    type="button"
                    onClick={() => setTeamSubmitted(false)}
                    className="text-xs text-blue-600 font-bold underline cursor-pointer"
                  >
                    Đăng ký lại hoặc chọn Chi bộ khác
                  </button>
                </div>
              </div>
            ) : (
              <>
                {/* Cho phép chọn Tháng và Năm đăng ký */}
                <div className="grid grid-cols-2 gap-3 bg-red-50/60 p-3 rounded-xl border border-red-200">
                  <div>
                    <label className="block text-xs font-bold text-red-900 mb-1">1. Đăng ký cho Tháng (*)</label>
                    <select
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(Number(e.target.value))}
                      className="w-full p-2 border border-red-300 rounded-lg text-sm font-bold text-red-800 bg-white"
                    >
                      {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => (
                        <option key={m} value={m}>Tháng {m}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-red-900 mb-1">Năm (*)</label>
                    <select
                      value={selectedYear}
                      onChange={(e) => setSelectedYear(Number(e.target.value))}
                      className="w-full p-2 border border-red-300 rounded-lg text-sm font-bold text-red-800 bg-white"
                    >
                      {[2025, 2026, 2027].map(y => (
                        <option key={y} value={y}>Năm {y}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">2. Chọn Chi bộ của đồng chí (*)</label>
                  <select
                    value={teamSelectChiBo}
                    onChange={(e) => setTeamSelectChiBo(Number(e.target.value))}
                    className="w-full p-2.5 border rounded-lg text-sm font-bold text-red-800 bg-white border-red-200"
                  >
                    {branchDetails.map(b => (
                      <option key={b.id} value={b.id}>{b.name} (Bí thư: {b.biThu} - {b.sl} ĐV)</option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">3. Ngày họp dự kiến (*)</label>
                    <input
                      type="date"
                      required
                      value={teamDate}
                      onChange={(e) => {
                        const val = e.target.value;
                        setTeamDate(val);
                        if (val) {
                          const [yyyy, mm, dd] = val.split('-');
                          const d = new Date(Number(yyyy), Number(mm) - 1, Number(dd));
                          const days = ['Chủ nhật', 'thứ Hai', 'thứ Ba', 'thứ Tư', 'thứ Năm', 'thứ Sáu', 'thứ Bảy'];
                          setTeamDayOfWeek(days[d.getDay()]);
                          
                          // Tự động điều chỉnh tháng và năm tương ứng nếu người dùng chọn ngày ở tháng khác
                          const m = Number(mm);
                          const y = Number(yyyy);
                          if (m !== selectedMonth) setSelectedMonth(m);
                          if (y !== selectedYear) setSelectedYear(y);
                        }
                      }}
                      className="w-full p-2.5 border rounded-lg text-sm bg-white font-semibold"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Thứ trong tuần</label>
                    <input
                      type="text"
                      readOnly
                      value={teamDayOfWeek}
                      className="w-full p-2.5 border rounded-lg text-sm bg-slate-100 font-bold text-slate-700"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">4. Địa điểm tổ chức (*)</label>
                  <input
                    type="text"
                    required
                    value={teamLocation}
                    onChange={(e) => setTeamLocation(e.target.value)}
                    placeholder="VD: Hội trường Đội Quản lý thị trường..."
                    className="w-full p-2.5 border rounded-lg text-sm bg-white"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">5. Ghi chú thêm (nếu có)</label>
                  <input
                    type="text"
                    value={teamNote}
                    onChange={(e) => setTeamNote(e.target.value)}
                    placeholder="VD: Họp lồng ghép sinh hoạt chuyên đề..."
                    className="w-full p-2.5 border rounded-lg text-xs bg-white"
                  />
                </div>

                <div className="p-3 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-800">
                  📌 <b>Lưu ý:</b> Đề nghị các Chi bộ hoàn thành đăng ký trước ngày 20 hàng tháng.
                </div>

                <button
                  type="submit"
                  disabled={isSubmittingOnline}
                  className="w-full py-3 rounded-xl bg-red-600 hover:bg-red-700 disabled:bg-red-400 text-white font-bold text-sm shadow-md flex items-center justify-center gap-2 cursor-pointer transition-all"
                >
                  {isSubmittingOnline ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Send className="w-4 h-4" />}
                  {isSubmittingOnline ? 'Đang gửi lên hệ thống...' : `Gửi Đăng Ký Lịch Họp Tháng ${selectedMonth}/${selectedYear}`}
                </button>
              </>
            )}

            <div className="flex justify-between items-center pt-2 text-xs">
              <a 
                href={`${window.location.pathname}?mode=report&month=${selectedMonth}&year=${selectedYear}${apiParamStr}`}
                className="text-blue-600 hover:underline font-semibold"
              >
                👉 Chuyển sang Báo cáo ĐGXL sau họp
              </a>
              <a 
                href={window.location.pathname}
                className="text-slate-500 hover:text-red-600 font-medium"
              >
                ← Quay lại Bảng Quản Trị
              </a>
            </div>
          </form>
        </div>
      </div>
    );
  }

  // GIAO DIỆN DÀNH CHO CÁC ĐỘI TỰ BÁO CÁO ĐGXL SAU HỌP (mode=report)
  if (isPublicReportMode) {
    const selectedBranchInfo = branchDetails.find(b => b.id === reportBranchId) || branchDetails[0];

    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="bg-white border border-slate-200 rounded-2xl w-full max-w-xl shadow-xl overflow-hidden">
          <div className="bg-blue-700 p-6 text-white text-center relative">
            <div className="w-12 h-12 mx-auto rounded-xl bg-yellow-400 text-blue-900 flex items-center justify-center font-black text-2xl shadow mb-2">
              ★
            </div>
            <h2 className="text-base font-bold uppercase tracking-wide">ĐẢNG ỦY CHI CỤC QUẢN LÝ THỊ TRƯỜNG</h2>
            <h3 className="text-lg font-black mt-1">BÁO CÁO ĐÁNH GIÁ, XẾP LOẠI SINH HOẠT CHI BỘ</h3>
            <p className="text-sm font-semibold text-yellow-300 mt-1">Theo Hướng dẫn 01-HD/TU • Tháng {selectedMonth} năm {selectedYear}</p>
          </div>

          <form onSubmit={handleTeamSubmitReport} className="p-6 space-y-4">
            {reportSubmitted ? (
              <div className="p-5 rounded-xl bg-emerald-50 border border-emerald-200 text-center space-y-3">
                <CheckCircle2 className="w-12 h-12 text-emerald-600 mx-auto" />
                <h4 className="font-bold text-emerald-800 text-base">Gửi báo cáo thành công Tháng {selectedMonth}/{selectedYear}!</h4>
                <p className="text-xs text-emerald-700 leading-relaxed">
                  Kết quả đánh giá xếp loại sinh hoạt của <b>{selectedBranchInfo.name}</b> (Điểm: <b>{reportScore}</b> - Xếp loại: <b>{reportClassification}</b>) đã được lưu vào hệ thống máy chủ Đảng ủy.
                </p>

                {/* Nút Xuất Báo Cáo / In Ngay Sau Khi Nhập Xong */}
                <div className="pt-2 flex flex-col sm:flex-row items-center justify-center gap-2">
                  <button
                    type="button"
                    onClick={handleExportDGXLWordDoc}
                    className="w-full sm:w-auto px-4 py-2.5 bg-blue-700 hover:bg-blue-800 text-white rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 shadow cursor-pointer"
                  >
                    <FileText className="w-4 h-4" />
                    Xuất File Word (.DOC) In Báo Cáo
                  </button>

                  <button
                    type="button"
                    onClick={handlePrintDGXL}
                    className="w-full sm:w-auto px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 shadow cursor-pointer"
                  >
                    <Printer className="w-4 h-4" />
                    In Báo Cáo Trình Ký (1 Trang)
                  </button>
                </div>

                <div className="pt-2">
                  <button
                    type="button"
                    onClick={() => setReportSubmitted(false)}
                    className="text-xs text-blue-600 font-bold underline cursor-pointer"
                  >
                    Báo cáo lại hoặc chọn Chi bộ khác
                  </button>
                </div>
              </div>
            ) : (
              <>
                {/* Cho phép chọn Tháng và Năm báo cáo */}
                <div className="grid grid-cols-2 gap-3 bg-blue-50/60 p-3 rounded-xl border border-blue-200">
                  <div>
                    <label className="block text-xs font-bold text-blue-900 mb-1">1. Báo cáo cho Tháng (*)</label>
                    <select
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(Number(e.target.value))}
                      className="w-full p-2 border border-blue-300 rounded-lg text-sm font-bold text-blue-800 bg-white"
                    >
                      {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => (
                        <option key={m} value={m}>Tháng {m}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-xs font-bold text-blue-900 mb-1">Năm (*)</label>
                    <select
                      value={selectedYear}
                      onChange={(e) => setSelectedYear(Number(e.target.value))}
                      className="w-full p-2 border border-blue-300 rounded-lg text-sm font-bold text-blue-800 bg-white"
                    >
                      {[2025, 2026, 2027].map(y => (
                        <option key={y} value={y}>Năm {y}</option>
                      ))}
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">2. Chọn Chi bộ báo cáo (*)</label>
                  <select
                    value={reportBranchId}
                    onChange={(e) => setReportBranchId(Number(e.target.value))}
                    className="w-full p-2.5 border rounded-lg text-sm font-bold text-blue-800 bg-white border-blue-200"
                  >
                    {branchDetails.map(b => (
                      <option key={b.id} value={b.id}>{b.name} (Bí thư: {b.biThu} - {b.sl} ĐV)</option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">3. Số đảng viên dự / Tổng số (*)</label>
                    <input
                      type="text"
                      required
                      value={reportAttendance}
                      onChange={(e) => setReportAttendance(e.target.value)}
                      placeholder="VD: 18/18 hoặc 9/9"
                      className="w-full p-2.5 border rounded-lg text-sm bg-white font-bold"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">4. Ngày họp thực tế (*)</label>
                    <input
                      type="date"
                      required
                      value={reportMeetingDate}
                      onChange={(e) => {
                        const val = e.target.value;
                        setReportMeetingDate(val);
                        if (val) {
                          const [yyyy, mm] = val.split('-');
                          const m = Number(mm);
                          const y = Number(yyyy);
                          if (m !== selectedMonth) setSelectedMonth(m);
                          if (y !== selectedYear) setSelectedYear(y);
                        }
                      }}
                      className="w-full p-2.5 border rounded-lg text-sm bg-white font-semibold"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">5. Điểm tự đánh giá (0-100) (*)</label>
                    <input
                      type="number"
                      required
                      min={0}
                      max={100}
                      value={reportScore}
                      onChange={(e) => {
                        const score = Number(e.target.value);
                        setReportScore(score);
                        if (score >= 90) setReportClassification('Tốt');
                        else if (score >= 70) setReportClassification('Khá');
                        else if (score >= 50) setReportClassification('Trung bình');
                        else setReportClassification('Kém');
                      }}
                      className="w-full p-2.5 border rounded-lg text-base font-black text-center text-blue-700 bg-white"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">6. Mức xếp loại chi bộ (*)</label>
                    <select
                      value={reportClassification}
                      onChange={(e) => setReportClassification(e.target.value)}
                      className="w-full p-2.5 border rounded-lg text-sm font-bold text-emerald-700 bg-emerald-50 border-emerald-200"
                    >
                      <option value="Tốt">Tốt (Từ 90-100 điểm)</option>
                      <option value="Khá">Khá (Từ 70-89 điểm)</option>
                      <option value="Trung bình">Trung bình (Từ 50-69 điểm)</option>
                      <option value="Kém">Kém (Dưới 50 điểm)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    7. Thuyết minh điểm trừ / Ghi chú biến động (nếu có)
                  </label>
                  <textarea
                    rows={3}
                    value={reportDeductionNote}
                    onChange={(e) => setReportDeductionNote(e.target.value)}
                    placeholder="VD: Có 01 đ/c vắng mặt có lý do (chuyển sinh hoạt tạm thời)..."
                    className="w-full p-2.5 border rounded-lg text-xs bg-white"
                  />
                </div>

                <div className="p-3 rounded-lg bg-blue-50 border border-blue-200 text-xs text-blue-800">
                  📌 <b>Theo HD 01-HD/TU:</b> Chi bộ đạt từ 90-100 điểm xếp loại <b>Tốt</b>; từ 70-89 điểm xếp loại <b>Khá</b>; từ 50-69 điểm xếp loại <b>Trung bình</b>; dưới 50 điểm xếp loại <b>Kém</b>.
                </div>

                <button
                  type="submit"
                  disabled={isSubmittingOnline}
                  className="w-full py-3 rounded-xl bg-blue-700 hover:bg-blue-800 disabled:bg-blue-400 text-white font-bold text-sm shadow-md flex items-center justify-center gap-2 cursor-pointer transition-all"
                >
                  {isSubmittingOnline ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Send className="w-4 h-4" />}
                  {isSubmittingOnline ? 'Đang gửi lên hệ thống...' : `Gửi Báo Cáo Đánh Giá Xếp Loại Tháng ${selectedMonth}/${selectedYear}`}
                </button>
              </>
            )}

            <div className="flex justify-between items-center pt-2 text-xs">
              <a 
                href={`${window.location.pathname}?mode=register&month=${selectedMonth}&year=${selectedYear}${apiParamStr}`}
                className="text-red-600 hover:underline font-semibold"
              >
                👉 Chuyển sang Đăng ký Lịch họp (Trước ngày 20)
              </a>
              <a 
                href={window.location.pathname}
                className="text-slate-500 hover:text-blue-600 font-medium"
              >
                ← Quay lại Bảng Quản Trị
              </a>
            </div>
          </form>
        </div>
      </div>
    );
  }

  // GIAO DIỆN QUẢN TRỊ CHÍNH
  return (
    <div className={`min-h-screen ${darkMode ? 'dark' : ''}`}>
      <div className="bg-[var(--bg-main)] text-[var(--text-main)] min-h-screen flex flex-col font-sans">
        
        {/* TOAST THÔNG BÁO LƯU */}
        {saveToast && (
          <div className="fixed top-4 right-4 z-50 bg-emerald-600 text-white px-4 py-3 rounded-xl shadow-2xl flex items-center gap-2 text-xs font-bold animate-bounce no-print">
            <CheckCircle2 className="w-4 h-4 text-emerald-200" />
            {saveToast}
          </div>
        )}

        {/* HEADER BAR */}
        <header className="border-b border-[var(--border-color)] bg-[var(--bg-card)] sticky top-0 z-30 px-6 py-3.5 no-print">
          <div className="max-w-7xl mx-auto flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-red-600 flex items-center justify-center text-yellow-300 font-black shadow-md shadow-red-500/20 text-xl">
                ★
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h1 className="text-base font-bold tracking-tight text-[var(--text-main)]">
                    ĐẢNG ỦY CHI CỤC QUẢN LÝ THỊ TRƯỜNG AN GIANG
                  </h1>
                  <span className="badge bg-red-50 text-red-700 border border-red-200 dark:bg-red-950/40 dark:text-red-300 dark:border-red-800">
                    Cán bộ tổng hợp
                  </span>
                </div>
                <p className="text-xs text-[var(--text-muted)]">Quy trình Đăng ký ngày 20, Duyệt ngày 25, Lưu thông báo, ĐGXL hằng tháng theo HD 01 & Báo cáo 06 Tháng</p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              {/* Nút Đồng bộ Google Sheets */}
              <button
                onClick={() => {
                  if (!googleScriptUrl) {
                    setIsApiModalOpen(true);
                  } else {
                    fetchFromGoogleSheets(googleScriptUrl);
                  }
                }}
                disabled={isSyncing}
                className={`inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg border transition-all cursor-pointer ${
                  googleScriptUrl 
                    ? 'bg-emerald-50 text-emerald-700 border-emerald-300 hover:bg-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-300' 
                    : 'bg-amber-50 text-amber-700 border-amber-300 hover:bg-amber-100 animate-pulse'
                }`}
                title={googleScriptUrl ? `Đồng bộ với Google Sheets (Lần cuối: ${lastSyncTime || 'Chưa đồng bộ'})` : 'Bấm để kết nối Google Sheets'}
              >
                <RefreshCw className={`w-3.5 h-3.5 ${isSyncing ? 'animate-spin' : ''}`} />
                {isSyncing ? 'Đang tải...' : googleScriptUrl ? 'Đồng bộ Sheets' : 'Kết nối Sheets'}
              </button>

              <button
                onClick={() => setIsApiModalOpen(true)}
                className="p-2 rounded-lg border border-[var(--border-color)] text-[var(--text-muted)] hover:text-[var(--text-main)] transition-all cursor-pointer"
                title="Cài đặt kết nối Google Sheets"
              >
                <Settings className="w-4 h-4" />
              </button>

              {activeTab === 'dgxl' ? (
                <>
                  <button 
                    onClick={handleExportDGXLWordDoc}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-blue-700 hover:bg-blue-800 text-white shadow-sm transition-all cursor-pointer"
                    title="Tải về file Word (.DOC) ĐGXL chuẩn 1 trang A4 y hệt mẫu"
                  >
                    <FileText className="w-4 h-4" />
                    Xuất File Word (.DOC)
                  </button>

                  <button 
                    onClick={handlePrintDGXL}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-red-600 hover:bg-red-700 text-white shadow-sm transition-all cursor-pointer"
                    title="In bảng ĐGXL trình ký Bí thư Đảng ủy (Chuẩn 1 trang A4)"
                  >
                    <Printer className="w-4 h-4" />
                    In ĐGXL (1 Trang)
                  </button>
                </>
              ) : (
                <>
                  <button 
                    onClick={handleExportWordDoc}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-blue-700 hover:bg-blue-800 text-white shadow-sm transition-all cursor-pointer"
                    title="Tải về file Word (.DOC) chuẩn 100% mở bằng Microsoft Word không bao giờ lỗi"
                  >
                    <FileText className="w-4 h-4" />
                    Tải File Word (.DOC)
                  </button>

                  <button 
                    onClick={handlePrintSchedule}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-red-600 hover:bg-red-700 text-white shadow-sm transition-all cursor-pointer"
                    title="In thông báo lịch sinh hoạt trình ký Bí thư Đảng ủy (Chuẩn khít 1 trang A4)"
                  >
                    <Printer className="w-4 h-4" />
                    In Thông Báo (1 Trang)
                  </button>
                </>
              )}

              <button 
                onClick={() => setIsShareModalOpen(true)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white shadow-sm transition-all cursor-pointer"
                title="Tạo link chia sẻ cho các Chi bộ (Đăng ký lịch họp & Báo cáo ĐGXL)"
              >
                <Share2 className="w-4 h-4" />
                Gửi Link Chi Bộ
              </button>

              <button 
                onClick={() => setDarkMode(!darkMode)}
                className="p-2 rounded-lg border border-[var(--border-color)] text-[var(--text-muted)] hover:text-[var(--text-main)] transition-all cursor-pointer"
              >
                {darkMode ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {/* MAIN TABS */}
          <div className="max-w-7xl mx-auto flex gap-6 mt-3 border-t border-[var(--border-color)] pt-2 overflow-x-auto">
            {[
              { id: 'meetings', label: '📅 1. Lịch Họp & Duyệt Trình Ký (Ngày 20-25)', icon: Calendar },
              { id: 'archive', label: '📁 2. Kho Lưu Trữ Thông Báo Lịch Họp', icon: Archive },
              { id: 'dgxl', label: '⭐ 3. Đánh Giá Xếp Loại Hằng Tháng (HD 01-HD/TU)', icon: FileCheck2 },
              { id: 'sixmonths', label: '📊 4. Biểu Tổng Hợp 06 Tháng / Năm', icon: BookOpen },
              { id: 'dashboard', label: '📈 5. Tổng quan Đảng bộ (157 ĐV)', icon: Sparkles },
              { id: 'members', label: '👥 6. Danh sách Đảng viên', icon: Users }
            ].map(tab => {
              const Icon = tab.icon;
              const isActive = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  className={`flex items-center gap-2 py-2 px-1 text-xs font-medium border-b-2 transition-colors whitespace-nowrap cursor-pointer ${
                    isActive 
                      ? 'border-red-600 text-red-600 dark:text-red-400 font-bold' 
                      : 'border-transparent text-[var(--text-muted)] hover:text-[var(--text-main)]'
                  }`}
                >
                  <Icon className="w-4 h-4" />
                  {tab.label}
                </button>
              );
            })}
          </div>
        </header>

        {/* PRINT DOCUMENT 1: THÔNG BÁO LỊCH HỌP (1 TRANG A4) */}
        {printDocType === 'schedule' && (
          <div className="print-only a4-page-standard">
            <table className="w-full border-none mb-1">
              <tbody>
                <tr>
                  <td className="w-[44%] text-center align-top border-none leading-tight" style={{ fontSize: '11pt' }}>
                    <div>ĐẢNG BỘ CỤC QUẢN LÝ THỊ TRƯỜNG</div>
                    <div style={{ fontWeight: 'bold' }}>ĐẢNG ỦY BỘ PHẬN</div>
                    <div style={{ marginTop: '-2px', marginBottom: '2px' }}>*</div>
                    <div>Số &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-TB/ĐUBP</div>
                  </td>
                  <td className="w-[56%] text-center align-top border-none leading-tight" style={{ fontSize: '12pt' }}>
                    <div style={{ fontWeight: 'bold' }}>ĐẢNG CỘNG SẢN VIỆT NAM</div>
                    <div style={{ height: '8px' }}></div>
                    <div style={{ fontStyle: 'italic' }}>
                      An Giang, ngày &nbsp;&nbsp;&nbsp; tháng {String(selectedMonth).padStart(2, '0')} năm {selectedYear}
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>

            <div className="text-center my-1.5 leading-tight">
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>THÔNG BÁO</div>
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>
                Lịch sinh hoạt lệ các chi bộ trực thuộc Tháng {selectedMonth} năm {selectedYear}
              </div>
              <div style={{ fontSize: '11pt' }}>-----</div>
            </div>

            <div className="text-justify indent-8 mb-1 leading-snug" style={{ fontSize: '13pt' }}>
              Căn cứ Quy chế làm việc của Đảng ủy bộ phận; Đảng ủy bộ phận thông báo lịch sinh hoạt lệ các chi bộ trực thuộc trong Tháng {selectedMonth} năm {selectedYear} cụ thể như sau:
            </div>

            <table className="w-full my-1 border border-black" style={{ fontSize: '11pt' }}>
              <thead>
                <tr className="bg-slate-100 text-center font-bold" style={{ height: '24px' }}>
                  <th className="border border-black px-1 py-0.5 w-[4%] text-center">STT</th>
                  <th className="border border-black px-1 py-0.5 w-[14%] text-left">Chi bộ</th>
                  <th className="border border-black px-1 py-0.5 w-[7%] text-center">Số lượng</th>
                  <th className="border border-black px-1 py-0.5 w-[22%] text-center">Thời gian sinh hoạt</th>
                  <th className="border border-black px-1 py-0.5 w-[15%] text-left">Địa điểm</th>
                  <th className="border border-black px-1 py-0.5 w-[22%] text-left">Bí thư, điện thoại</th>
                  <th className="border border-black px-1 py-0.5 w-[16%] text-left">Ghi chú</th>
                </tr>
              </thead>
              <tbody>
                {currentMonthSchedules.map((item, idx) => (
                  <tr key={item.chiBoId} style={{ height: '20px' }}>
                    <td className="border border-black px-1 py-0.5 text-center">{idx + 1}</td>
                    <td className="border border-black px-1 py-0.5">{item.chiBo}</td>
                    <td className="border border-black px-1 py-0.5 text-center">{item.sl}</td>
                    <td className="border border-black px-1 py-0.5 text-center">{item.thoiGian || ''}</td>
                    <td className="border border-black px-1 py-0.5">{item.diaDiem || ''}</td>
                    <td className="border border-black px-1 py-0.5">{item.biThu} - {item.sdt}</td>
                    <td className="border border-black px-1 py-0.5">{item.ghiChu || ''}</td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="text-justify indent-8 mb-2 leading-snug" style={{ fontSize: '13pt' }}>
              Đề nghị các đồng chí Đảng ủy viên phụ trách các chi bộ sắp xếp thời gian cùng dự họp với các chi bộ theo thời gian và địa điểm nêu trên./.
            </div>

            <table className="w-full border-none mt-2">
              <tbody>
                <tr>
                  <td className="w-[50%] align-top border-none leading-tight" style={{ fontSize: '10.5pt' }}>
                    <div style={{ fontStyle: 'italic', fontWeight: 'bold' }}>Nơi nhận:</div>
                    <div>- Các ĐUV (dự);</div>
                    <div>- 13 chi bộ trực thuộc (t/h);</div>
                    <div>- Lưu: ĐUBP.</div>
                  </td>
                  <td className="w-[50%] text-center align-top border-none leading-tight" style={{ fontSize: '12pt' }}>
                    <div style={{ fontWeight: 'bold' }}>T/M ĐẢNG ỦY BỘ PHẬN</div>
                    <div style={{ fontWeight: 'normal' }}>BÍ THƯ</div>
                    <div style={{ height: '40px' }}></div>
                    <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>Nguyễn Trung Tiến</div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* PRINT DOCUMENT 2: BIỂU ĐGXL THEO HD 01-HD/TU (CHUẨN 10 CỘT 1 TRANG A4) */}
        {printDocType === 'dgxl' && (
          <div className="print-only a4-page-standard">
            <table className="w-full border-none mb-1">
              <tbody>
                <tr>
                  <td className="w-[54%] text-center align-top border-none leading-tight" style={{ fontSize: '11pt' }}>
                    <div>ĐẢNG BỘ SỞ CÔNG THƯƠNG TỈNH AN GIANG</div>
                    <div style={{ fontWeight: 'bold' }}>ĐẢNG ỦY CHI CỤC QUẢN LÝ THỊ TRƯỜNG</div>
                    <div style={{ marginTop: '-2px', marginBottom: '2px' }}>*</div>
                  </td>
                  <td className="w-[46%] text-center align-top border-none leading-tight" style={{ fontSize: '12pt' }}>
                    <div style={{ fontWeight: 'bold' }}>ĐẢNG CỘNG SẢN VIỆT NAM</div>
                    <div style={{ height: '6px' }}></div>
                    <div style={{ fontStyle: 'italic' }}>
                      An Giang, ngày 10 tháng {String(selectedMonth).padStart(2, '0')} năm {selectedYear}
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>

            <div className="text-center my-1.5 leading-tight">
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>DANH SÁCH</div>
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>
                TỔNG HỢP ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG SINH HOẠT CHI BỘ
              </div>
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>
                Tháng {selectedMonth} NĂM {selectedYear}
              </div>
              <div style={{ fontSize: '11pt' }}>-----</div>
            </div>

            <table className="w-full my-1 border border-black" style={{ fontSize: '10.5pt' }}>
              <thead>
                <tr className="bg-slate-100 text-center font-bold">
                  <th rowSpan={2} className="border border-black px-1 py-0.5 w-[4%] text-center">S<br/>T<br/>T</th>
                  <th rowSpan={2} className="border border-black px-1.5 py-0.5 w-[16%] text-left">Tên đơn vị</th>
                  <th rowSpan={2} className="border border-black px-1 py-0.5 w-[9%] text-center">Số<br/>lượng<br/>đảng viên</th>
                  <th rowSpan={2} className="border border-black px-1 py-0.5 w-[10%] text-center">Ngày<br/>họp<br/>chi<br/>bộ</th>
                  <th colSpan={5} className="border border-black px-1 py-0.5 text-center">
                    Mức đánh giá xếp loại sau buổi<br/>sinh hoạt chi bộ hằng tháng
                  </th>
                  <th rowSpan={2} className="border border-black px-1.5 py-0.5 w-[23%] text-left">
                    Ghi chú<br/>(thuyết minh điểm trừ)
                  </th>
                </tr>
                <tr className="bg-slate-100 text-center font-bold" style={{ fontSize: '10pt' }}>
                  <th className="border border-black px-1 py-0.5 w-[9%] text-center">KQ số<br/>điểm<br/>được<br/>đánh giá</th>
                  <th className="border border-black px-1 py-0.5 w-[7%] text-center">Tốt</th>
                  <th className="border border-black px-1 py-0.5 w-[7%] text-center">Khá</th>
                  <th className="border border-black px-1 py-0.5 w-[7%] text-center">Trung<br/>bình</th>
                  <th className="border border-black px-1 py-0.5 w-[7%] text-center">Kém</th>
                </tr>
                <tr className="text-center italic bg-slate-50" style={{ fontSize: '9pt' }}>
                  <td className="border border-black py-0.5">1</td>
                  <td className="border border-black py-0.5">2</td>
                  <td className="border border-black py-0.5">3</td>
                  <td className="border border-black py-0.5">4</td>
                  <td className="border border-black py-0.5">5</td>
                  <td className="border border-black py-0.5">6</td>
                  <td className="border border-black py-0.5">7</td>
                  <td className="border border-black py-0.5">8</td>
                  <td className="border border-black py-0.5">9</td>
                  <td className="border border-black py-0.5">10</td>
                </tr>
              </thead>
              <tbody>
                {currentMonthDGXL.map((item, idx) => (
                  <tr key={item.chiBoId} style={{ height: '19px' }}>
                    <td className="border border-black px-1 py-0.5 text-center">{idx + 1}</td>
                    <td className="border border-black px-1.5 py-0.5">{item.chiBo}</td>
                    <td className="border border-black px-1 py-0.5 text-center">{item.sl || ''}</td>
                    <td className="border border-black px-1 py-0.5 text-center">{item.ngayHop || ''}</td>
                    <td className="border border-black px-1 py-0.5 text-center">{item.diemDG !== '' ? item.diemDG : ''}</td>
                    <td className="border border-black px-1 py-0.5 text-center font-bold">{item.mucXepLoai === 'Tốt' ? 'x' : ''}</td>
                    <td className="border border-black px-1 py-0.5 text-center font-bold">{item.mucXepLoai === 'Khá' ? 'x' : ''}</td>
                    <td className="border border-black px-1 py-0.5 text-center font-bold">{item.mucXepLoai === 'Trung bình' ? 'x' : ''}</td>
                    <td className="border border-black px-1 py-0.5 text-center font-bold">{item.mucXepLoai === 'Kém' ? 'x' : ''}</td>
                    <td className="border border-black px-1.5 py-0.5" style={{ fontSize: '9.5pt' }}>{item.ghiChuTruDiem || ''}</td>
                  </tr>
                ))}
                <tr className="font-bold text-center" style={{ height: '20px' }}>
                  <td colSpan={2} className="border border-black px-1.5 py-0.5 text-center">Tổng cộng</td>
                  <td className="border border-black px-1 py-0.5"></td>
                  <td className="border border-black px-1 py-0.5"></td>
                  <td className="border border-black px-1 py-0.5"></td>
                  <td className="border border-black px-1 py-0.5">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Tốt').length || ''}</td>
                  <td className="border border-black px-1 py-0.5">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Khá').length || ''}</td>
                  <td className="border border-black px-1 py-0.5">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Trung bình').length || ''}</td>
                  <td className="border border-black px-1 py-0.5">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Kém').length || ''}</td>
                  <td className="border border-black px-1.5 py-0.5"></td>
                </tr>
              </tbody>
            </table>

            <div className="text-justify my-1 leading-snug" style={{ fontSize: '10.5pt' }}>
              <b>* Ghi chú:</b> {currentMonthNote}
            </div>

            <table className="w-full border-none mt-3">
              <tbody>
                <tr>
                  <td className="w-[50%] text-center align-top border-none leading-tight" style={{ fontSize: '12pt' }}>
                    <div style={{ fontWeight: 'bold' }}>NGƯỜI LẬP BIỂU</div>
                    <div style={{ height: '45px' }}></div>
                    <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>Trần Đình Chinh</div>
                  </td>
                  <td className="w-[50%] text-center align-top border-none leading-tight" style={{ fontSize: '12pt' }}>
                    <div style={{ fontWeight: 'bold' }}>T/M ĐẢNG ỦY</div>
                    <div style={{ fontWeight: 'normal' }}>BÍ THƯ</div>
                    <div style={{ height: '40px' }}></div>
                    <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>Nguyễn Trung Tiến</div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* MAIN BODY CONTENT */}
        <main className="flex-1 max-w-7xl w-full mx-auto p-6 space-y-6 no-print">

          {/* TAB 1: LỊCH HỌP & DUYỆT TRÌNH KÝ */}
          {activeTab === 'meetings' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-base font-bold text-[var(--text-main)]">
                      Lịch Sinh Hoạt Lệ 13 Chi Bộ Trực Thuộc Tháng {selectedMonth}/{selectedYear}
                    </h3>
                    <span className="badge bg-blue-50 text-blue-700 border border-blue-200">
                      Bí thư: Nguyễn Trung Tiến
                    </span>
                  </div>
                  <p className="text-xs text-[var(--text-muted)] mt-0.5">
                    💡 <b>Quyền Quản trị:</b> Đồng chí có thể sửa ngày/địa điểm trực tiếp trên bảng, nhấp biểu tượng 📅 chọn ngày từ lịch, hoặc bấm <b>"✏️ Sửa"</b> để mở popup. Mọi thay đổi sẽ tự động lưu và đồng bộ!
                  </p>
                </div>

                <div className="flex items-center gap-2.5 flex-wrap">
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs font-semibold text-[var(--text-muted)]">Tháng:</span>
                    <select 
                      value={selectedMonth} 
                      onChange={(e) => setSelectedMonth(Number(e.target.value))}
                      className="w-24 text-xs font-bold"
                    >
                      {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <option key={m} value={m}>Tháng {m}</option>)}
                    </select>
                  </div>

                  <div className="flex items-center gap-1.5">
                    <span className="text-xs font-semibold text-[var(--text-muted)]">Năm:</span>
                    <select 
                      value={selectedYear} 
                      onChange={(e) => setSelectedYear(Number(e.target.value))}
                      className="w-24 text-xs font-bold"
                    >
                      {[2025, 2026, 2027].map(y => <option key={y} value={y}>Năm {y}</option>)}
                    </select>
                  </div>

                  <button
                    onClick={handleSaveAllMeetingsToGoogleSheets}
                    disabled={isSyncing}
                    className="px-3.5 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs shadow-sm flex items-center gap-1.5 cursor-pointer transition-all"
                    title="Lưu toàn bộ danh sách Lịch họp tháng này lên Google Sheets"
                  >
                    <Save className="w-4 h-4" />
                    {isSyncing ? 'Đang lưu...' : 'Lưu Lên Google Sheets'}
                  </button>

                  <button
                    onClick={handleApproveAndArchive}
                    className="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-sm flex items-center gap-1.5 cursor-pointer"
                  >
                    <CheckCheck className="w-4 h-4" />
                    Lãnh Đạo Thống Nhất & Lưu File
                  </button>
                </div>
              </div>

              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th className="w-10 text-center">STT</th>
                        <th className="w-44">Chi bộ</th>
                        <th className="w-16 text-center">Số lượng</th>
                        <th className="w-64">Thời gian sinh hoạt (Sửa / Chọn lịch)</th>
                        <th className="w-44">Địa điểm</th>
                        <th className="w-44">Bí thư, điện thoại</th>
                        <th>Ghi chú biến động</th>
                        <th className="w-32 text-center">Trạng thái</th>
                        <th className="w-16 text-center">Thao tác</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentMonthSchedules.map((item, idx) => (
                        <tr key={item.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{item.chiBo}</td>
                          <td className="text-center font-bold text-red-600 text-xs">{item.sl}</td>
                          <td>
                            <div className="flex items-center gap-1.5">
                              <input
                                type="text"
                                value={item.thoiGian}
                                onChange={(e) => handleUpdateSchedule(item.chiBoId, 'thoiGian', e.target.value)}
                                placeholder="VD: 03/10/2026 (thứ Bảy)"
                                className="text-xs font-medium text-blue-600 dark:text-blue-400 flex-1"
                              />
                              {/* DatePicker Icon Button */}
                              <div className="relative inline-block" title="Chọn ngày từ lịch (tự tính thứ)">
                                <input
                                  type="date"
                                  onChange={(e) => handlePickDateForMeeting(item.chiBoId, e.target.value)}
                                  className="opacity-0 absolute inset-0 w-full h-full cursor-pointer z-10"
                                />
                                <div className="p-1 rounded bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 text-slate-600 dark:text-slate-300 cursor-pointer">
                                  <CalendarDays className="w-3.5 h-3.5 text-blue-600" />
                                </div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <input
                              type="text"
                              value={item.diaDiem}
                              onChange={(e) => handleUpdateSchedule(item.chiBoId, 'diaDiem', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <div className="font-bold text-xs">{item.biThu}</div>
                            <div className="text-xs text-[var(--text-muted)] font-mono">{item.sdt}</div>
                          </td>
                          <td>
                            <input
                              type="text"
                              placeholder="Thêm ghi chú..."
                              value={item.ghiChu || ''}
                              onChange={(e) => handleUpdateSchedule(item.chiBoId, 'ghiChu', e.target.value)}
                              className="text-xs text-amber-700 dark:text-amber-300"
                            />
                          </td>
                          <td className="text-center">
                            <span className={`badge ${
                              item.trangThai === 'Đã thống nhất' ? 'bg-emerald-50 text-emerald-700 border border-emerald-200 dark:bg-emerald-950/40 dark:text-emerald-300' :
                              item.trangThai === 'Đã đăng ký (Chờ duyệt)' ? 'bg-blue-50 text-blue-700 border border-blue-200' :
                              'bg-amber-50 text-amber-700 border border-amber-200'
                            }`}>
                              {item.trangThai}
                            </span>
                          </td>
                          <td className="text-center">
                            <button
                              onClick={() => handleOpenAdminEditModal('meeting', item.chiBoId)}
                              className="px-2 py-1 bg-slate-100 hover:bg-blue-50 text-blue-700 border border-slate-200 rounded text-xs font-semibold inline-flex items-center gap-1 cursor-pointer"
                              title="Mở popup chỉnh sửa chi tiết"
                            >
                              <Edit className="w-3 h-3" />
                              Sửa
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          )}

          {/* TAB 2: KHO LƯU TRỮ THÔNG BÁO LỊCH HỌP ĐÃ BAN HÀNH */}
          {activeTab === 'archive' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex items-center justify-between">
                <div>
                  <h3 className="text-base font-bold text-[var(--text-main)]">Kho Lưu Trữ Thông Báo Lịch Họp (Theo Từng Tháng)</h3>
                  <p className="text-xs text-[var(--text-muted)]">Nơi lưu lại các Thông báo lịch sinh hoạt chi bộ đã ban hành để sau này tra cứu, đối chiếu</p>
                </div>
              </div>

              {archivedNotices.length === 0 ? (
                <div className="card-glass p-8 text-center text-[var(--text-muted)] text-sm">
                  Chưa có thông báo nào được lưu. Khi đồng chí bấm <b>"Lãnh đạo Thống nhất & Lưu File"</b> tại Tab Lịch Họp, hệ thống sẽ tự động lưu vào đây!
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {archivedNotices.map((arc, idx) => (
                    <div key={arc.id} className="card-glass p-5 border-t-4 border-red-500 space-y-3">
                      <div className="flex items-center justify-between">
                        <span className="badge bg-red-50 text-red-700 font-bold border border-red-200 text-xs">
                          Thông Báo Tháng {arc.thang}/{arc.nam}
                        </span>
                        <span className="text-xs text-[var(--text-muted)]">Lưu ngày: {arc.ngayBanHanh}</span>
                      </div>

                      <h4 className="font-bold text-sm text-[var(--text-main)]">
                        Lịch sinh hoạt lệ 13 Chi bộ trực thuộc Tháng {arc.thang}/{arc.nam}
                      </h4>
                      <p className="text-xs text-[var(--text-muted)]">
                        Tổng số đảng viên: <b>{arc.tongSoDV} ĐV</b> • Người ký: <b>{arc.nguoiKy}</b>
                      </p>

                      <div className="border border-[var(--border-color)] rounded-lg p-2 max-h-48 overflow-y-auto bg-slate-50 dark:bg-slate-900/50 text-xs space-y-1">
                        {arc.chiTiet.map(c => (
                          <div key={c.chiBoId} className="flex items-center justify-between border-b border-[var(--border-color)]/50 pb-1">
                            <span className="font-semibold">{c.chiBo} ({c.sl} ĐV):</span>
                            <span className="text-blue-600 dark:text-blue-400">{c.thoiGian}</span>
                          </div>
                        ))}
                      </div>

                      <div className="flex items-center justify-end gap-2 pt-2">
                        <button
                          onClick={() => {
                            setSelectedMonth(arc.thang);
                            setSelectedYear(arc.nam);
                            setActiveTab('meetings');
                          }}
                          className="px-3 py-1.5 rounded bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 text-xs font-semibold cursor-pointer"
                        >
                          Mở lại Tháng {arc.thang}
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* TAB 3: ĐÁNH GIÁ XẾP LOẠI CHẤT LƯỢNG SINH HOẠT HẰNG THÁNG (ĐGXL THEO HD 01-HD/TU) */}
          {activeTab === 'dgxl' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-base font-bold text-[var(--text-main)]">
                      DANH SÁCH TỔNG HỢP ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG SINH HOẠT CHI BỘ
                    </h3>
                    <span className="badge bg-emerald-50 text-emerald-700 border border-emerald-200 font-bold">
                      Tháng {selectedMonth} năm {selectedYear}
                    </span>
                  </div>
                  <p className="text-xs text-[var(--text-muted)] mt-1">
                    Bảng 10 cột chuẩn theo Hướng dẫn 01-HD/TU. Quản trị viên có thể chỉnh sửa ngày họp, điểm số, xếp loại và bấm <b>"Lưu Lên Google Sheets"</b> để đồng bộ dữ liệu.
                  </p>
                </div>

                <div className="flex items-center gap-2.5 flex-wrap">
                  <select 
                    value={selectedMonth} 
                    onChange={(e) => setSelectedMonth(Number(e.target.value))}
                    className="w-24 text-xs font-bold"
                  >
                    {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <option key={m} value={m}>Tháng {m}</option>)}
                  </select>

                  <select 
                    value={selectedYear} 
                    onChange={(e) => setSelectedYear(Number(e.target.value))}
                    className="w-24 text-xs font-bold"
                  >
                    {[2025, 2026, 2027].map(y => <option key={y} value={y}>Năm {y}</option>)}
                  </select>

                  <button 
                    onClick={handleSaveAllDGXLToGoogleSheets}
                    disabled={isSyncing}
                    className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all"
                    title="Lưu toàn bộ bảng ĐGXL lên Google Sheets"
                  >
                    <Save className="w-4 h-4" /> {isSyncing ? 'Đang lưu...' : 'Lưu Lên Sheets'}
                  </button>

                  <button 
                    onClick={handleExportDGXLWordDoc}
                    className="px-3.5 py-2 bg-blue-700 hover:bg-blue-800 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all"
                  >
                    <FileText className="w-4 h-4" /> Xuất Word (.DOC)
                  </button>

                  <button 
                    onClick={handlePrintDGXL}
                    className="px-3.5 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all"
                  >
                    <Printer className="w-4 h-4" /> In Báo Cáo (1 Trang)
                  </button>

                  <button 
                    onClick={handleExportDGXLToExcel}
                    className="px-3.5 py-2 bg-slate-700 hover:bg-slate-800 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" /> Xuất Excel
                  </button>
                </div>
              </div>

              {/* BẢNG 10 CỘT THEO HƯỚNG DẪN 01-HD/TU (Y HỆT HÌNH ẢNH MẪU) */}
              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th rowSpan={2} className="w-10 text-center">S<br/>T<br/>T</th>
                        <th rowSpan={2} className="w-40">Tên đơn vị</th>
                        <th rowSpan={2} className="w-28 text-center">Số<br/>lượng<br/>đảng viên</th>
                        <th rowSpan={2} className="w-36 text-center">Ngày<br/>họp<br/>chi<br/>bộ</th>
                        <th colSpan={5} className="text-center bg-blue-50/50 dark:bg-blue-950/30">
                          Mức đánh giá xếp loại sau buổi sinh hoạt chi bộ hằng tháng
                        </th>
                        <th rowSpan={2} className="w-64">Ghi chú<br/>(thuyết minh điểm trừ)</th>
                        <th rowSpan={2} className="w-16 text-center">Thao tác</th>
                      </tr>
                      <tr>
                        <th className="w-24 text-center">KQ số điểm<br/>được<br/>đánh giá</th>
                        <th className="w-14 text-center text-emerald-700 dark:text-emerald-400">Tốt</th>
                        <th className="w-14 text-center text-blue-700 dark:text-blue-400">Khá</th>
                        <th className="w-14 text-center text-amber-700 dark:text-amber-400">Trung<br/>bình</th>
                        <th className="w-14 text-center text-red-700 dark:text-red-400">Kém</th>
                      </tr>
                      <tr className="text-center text-[10px] text-[var(--text-muted)] bg-slate-50 dark:bg-slate-900/50 italic">
                        <th>1</th>
                        <th>2</th>
                        <th>3</th>
                        <th>4</th>
                        <th>5</th>
                        <th>6</th>
                        <th>7</th>
                        <th>8</th>
                        <th>9</th>
                        <th>10</th>
                        <th>-</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentMonthDGXL.map((d, idx) => (
                        <tr key={d.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{d.chiBo}</td>
                          <td className="text-center">
                            <input
                              type="text"
                              value={d.sl || ''}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'sl', e.target.value)}
                              placeholder="VD: 18/18"
                              className="w-16 text-center font-bold text-xs"
                            />
                          </td>
                          <td className="text-center">
                            <div className="flex items-center justify-center gap-1">
                              <input
                                type="text"
                                value={d.ngayHop || ''}
                                onChange={(e) => handleUpdateDGXL(d.chiBoId, 'ngayHop', e.target.value)}
                                placeholder="03/10/2026"
                                className="w-24 text-center text-xs font-medium text-blue-600 dark:text-blue-400"
                              />
                              <div className="relative inline-block" title="Chọn ngày từ lịch">
                                <input
                                  type="date"
                                  onChange={(e) => handlePickDateForDGXL(d.chiBoId, e.target.value)}
                                  className="opacity-0 absolute inset-0 w-full h-full cursor-pointer z-10"
                                />
                                <div className="p-1 rounded bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 text-slate-600 dark:text-slate-300 cursor-pointer">
                                  <CalendarDays className="w-3.5 h-3.5 text-blue-600" />
                                </div>
                              </div>
                            </div>
                          </td>
                          <td className="text-center">
                            <input
                              type="number"
                              min={0}
                              max={100}
                              value={d.diemDG !== '' ? d.diemDG : ''}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'diemDG', e.target.value === '' ? '' : Number(e.target.value))}
                              placeholder="100"
                              className="w-16 text-center font-black text-xs text-blue-700 dark:text-blue-300"
                            />
                          </td>
                          <td className="text-center font-black text-emerald-600">
                            {d.mucXepLoai === 'Tốt' ? 'x' : ''}
                          </td>
                          <td className="text-center font-black text-blue-600">
                            {d.mucXepLoai === 'Khá' ? 'x' : ''}
                          </td>
                          <td className="text-center font-black text-amber-600">
                            {d.mucXepLoai === 'Trung bình' ? 'x' : ''}
                          </td>
                          <td className="text-center font-black text-red-600">
                            {d.mucXepLoai === 'Kém' ? 'x' : ''}
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.ghiChuTruDiem || ''}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'ghiChuTruDiem', e.target.value)}
                              placeholder="Nhập thuyết minh nếu có..."
                              className="text-xs"
                            />
                          </td>
                          <td className="text-center">
                            <button
                              onClick={() => handleOpenAdminEditModal('dgxl', d.chiBoId)}
                              className="px-2 py-1 bg-slate-100 hover:bg-blue-50 text-blue-700 border border-slate-200 rounded text-xs font-semibold inline-flex items-center gap-1 cursor-pointer"
                              title="Mở popup chỉnh sửa chi tiết"
                            >
                              <Edit className="w-3 h-3" />
                              Sửa
                            </button>
                          </td>
                        </tr>
                      ))}
                      <tr className="font-bold bg-slate-50/80 dark:bg-slate-900/50 text-xs">
                        <td colSpan={2} className="text-center py-2.5">Tổng cộng</td>
                        <td className="text-center text-red-600 font-bold">157</td>
                        <td></td>
                        <td></td>
                        <td className="text-center text-emerald-600 font-black">
                          {currentMonthDGXL.filter(d => d.mucXepLoai === 'Tốt').length}
                        </td>
                        <td className="text-center text-blue-600 font-black">
                          {currentMonthDGXL.filter(d => d.mucXepLoai === 'Khá').length}
                        </td>
                        <td className="text-center text-amber-600 font-black">
                          {currentMonthDGXL.filter(d => d.mucXepLoai === 'Trung bình').length}
                        </td>
                        <td className="text-center text-red-600 font-black">
                          {currentMonthDGXL.filter(d => d.mucXepLoai === 'Kém').length}
                        </td>
                        <td></td>
                        <td></td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Ô NHẬP GHI CHÚ CHÂN TRANG ĐGXL */}
              <div className="card-glass p-4 space-y-2">
                <label className="block text-xs font-bold text-[var(--text-main)]">
                  * Ghi chú chân trang danh sách (Biến động tổng thể của Đảng bộ):
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={currentMonthNote}
                    onChange={(e) => handleUpdateOverallNote(e.target.value)}
                    className="text-xs flex-1"
                    placeholder="Nhập ghi chú tổng thể biến động đảng viên..."
                  />
                  <button
                    onClick={() => alert('Đã lưu ghi chú chân trang thành công!')}
                    className="px-4 py-2 bg-slate-200 dark:bg-slate-800 text-xs font-semibold rounded-lg hover:bg-slate-300 cursor-pointer"
                  >
                    Lưu ghi chú
                  </button>
                </div>
              </div>

            </div>
          )}

          {/* TAB 4: BIỂU TỔNG HỢP SỐ LIỆU 06 THÁNG / NĂM */}
          {activeTab === 'sixmonths' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <div className="flex items-center gap-2">
                    <h3 className="text-base font-bold text-[var(--text-main)]">
                      BIỂU TỔNG HỢP SỐ LIỆU BÁO CÁO 06 THÁNG ĐẦU NĂM {selectedYear}
                    </h3>
                    <span className="badge bg-purple-50 text-purple-700 border border-purple-200">
                      Đảng bộ Sở Công Thương
                    </span>
                  </div>
                  <p className="text-xs text-[var(--text-muted)] mt-0.5">
                    Phục vụ báo cáo sơ kết 6 tháng, tổng kết năm đầy đủ 12 cột chuyên đề theo mẫu quy định
                  </p>
                </div>

                <div className="flex items-center gap-3">
                  <select 
                    value={selectedYear} 
                    onChange={(e) => setSelectedYear(Number(e.target.value))}
                    className="w-28 text-xs font-bold"
                  >
                    {[2025, 2026, 2027].map(y => <option key={y} value={y}>Năm {y}</option>)}
                  </select>

                  <button 
                    onClick={() => {
                      const ws = XLSX.utils.json_to_sheet(sixMonthsData);
                      const wb = XLSX.utils.book_new();
                      XLSX.utils.book_append_sheet(wb, ws, `BaoCao_6Thang_${selectedYear}`);
                      XLSX.writeFile(wb, `Bieu_Tong_Hop_6_Thang_${selectedYear}.xlsx`);
                    }}
                    className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" /> Xuất Excel 6 Tháng
                  </button>
                </div>
              </div>

              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th className="w-10 text-center">STT</th>
                        <th className="w-36">Tên tổ chức đảng</th>
                        <th className="w-20 text-center">Đảng số</th>
                        <th className="w-28 text-center">Số lượng ĐV dự họp</th>
                        <th className="w-36">Phát biểu ý kiến</th>
                        <th className="w-40">Mô hình Dân vận khéo</th>
                        <th className="w-44">Học tập chuyên đề Bác Hồ, Bác Tôn</th>
                        <th className="w-40">Đọc bài viết về Bác</th>
                        <th className="w-44">Kế hoạch sinh hoạt chuyên đề</th>
                        <th className="w-20 text-center">Kết nạp đảng</th>
                        <th className="w-36">Kiểm tra, giám sát</th>
                        <th>Ghi chú</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sixMonthsData.map((row, idx) => (
                        <tr key={row.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{row.chiBo}</td>
                          <td className="text-center font-bold text-red-600 text-xs">{row.dangSo}</td>
                          <td>
                            <input
                              type="text"
                              value={row.soLgDuHop}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'soLgDuHop', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.phatBieuYkien}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'phatBieuYkien', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.danVanKheo}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'danVanKheo', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.hocTapChuyenDe}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'hocTapChuyenDe', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.docBaiViet}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'docBaiViet', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.khSinhHoatCD}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'khSinhHoatCD', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.knd}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'knd', e.target.value)}
                              className="text-xs text-center font-bold"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.ktgs}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'ktgs', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={row.ghiChu || ''}
                              onChange={(e) => handleUpdateSixMonths(row.chiBoId, 'ghiChu', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* TAB 5: TỔNG QUAN ĐẢNG BỘ */}
          {activeTab === 'dashboard' && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="card-glass p-5 flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-red-50 dark:bg-red-950/40 text-red-600 flex items-center justify-center font-black text-xl">
                    ★
                  </div>
                  <div>
                    <p className="text-xs text-[var(--text-muted)] font-medium">Tổng số đảng viên</p>
                    <h3 className="text-2xl font-black text-[var(--text-main)] mt-0.5">157 <span className="text-xs font-normal text-[var(--text-muted)]">đồng chí</span></h3>
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-blue-50 dark:bg-blue-950/40 text-blue-600 flex items-center justify-center font-black text-xl">
                    🏢
                  </div>
                  <div>
                    <p className="text-xs text-[var(--text-muted)] font-medium">Số chi bộ trực thuộc</p>
                    <h3 className="text-2xl font-black text-[var(--text-main)] mt-0.5">13 <span className="text-xs font-normal text-[var(--text-muted)]">chi bộ</span></h3>
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-emerald-50 dark:bg-emerald-950/40 text-emerald-600 flex items-center justify-center font-black text-xl">
                    ⭐
                  </div>
                  <div>
                    <p className="text-xs text-[var(--text-muted)] font-medium">Xếp loại Tốt Tháng {selectedMonth}</p>
                    <h3 className="text-2xl font-black text-[var(--text-main)] mt-0.5">
                      {currentMonthDGXL.filter(d => d.mucXepLoai === 'Tốt').length} / 13
                    </h3>
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-amber-50 dark:bg-amber-950/40 text-amber-600 flex items-center justify-center font-black text-xl">
                    📅
                  </div>
                  <div>
                    <p className="text-xs text-[var(--text-muted)] font-medium">Đã đăng ký lịch họp</p>
                    <h3 className="text-2xl font-black text-[var(--text-main)] mt-0.5">
                      {currentMonthSchedules.filter(s => s.thoiGian && s.thoiGian !== 'Chưa đăng ký').length} / 13
                    </h3>
                  </div>
                </div>
              </div>

              {/* Danh sách 13 Chi bộ trực thuộc */}
              <div className="card-glass p-5 space-y-4">
                <h3 className="text-base font-bold text-[var(--text-main)]">Danh Sách 13 Chi Bộ Trực Thuộc Đảng Bộ QLTT</h3>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                  {branchDetails.map(b => (
                    <div key={b.id} className="p-3.5 rounded-xl border border-[var(--border-color)] bg-slate-50/50 dark:bg-slate-900/30 space-y-1.5">
                      <div className="flex items-center justify-between">
                        <span className="font-bold text-sm text-[var(--text-main)]">{b.name}</span>
                        <span className="badge bg-red-50 text-red-700 font-bold border border-red-200">{b.sl} ĐV</span>
                      </div>
                      <div className="text-xs text-[var(--text-muted)]">
                        Bí thư: <b className="text-[var(--text-main)]">{b.biThu}</b> ({b.sdt})
                      </div>
                      <div className="text-xs text-[var(--text-muted)] truncate">
                        📍 {b.diaDiem}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB 6: DANH SÁCH ĐẢNG VIÊN */}
          {activeTab === 'members' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex flex-wrap items-center justify-between gap-4">
                <div className="flex items-center gap-3 flex-1 max-w-md">
                  <div className="relative w-full">
                    <Search className="w-4 h-4 absolute left-3 top-2.5 text-[var(--text-muted)]" />
                    <input
                      type="text"
                      placeholder="Tìm kiếm họ tên, số thẻ, chi bộ..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="pl-9 text-xs"
                    />
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  <select
                    value={filterChiBo}
                    onChange={(e) => setFilterChiBo(e.target.value)}
                    className="text-xs"
                  >
                    <option value="Tất cả">Tất cả chi bộ</option>
                    {chiBoList.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>

              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th className="w-10 text-center">STT</th>
                        <th className="w-48">Họ và tên</th>
                        <th className="w-32 text-center">Số thẻ ĐV</th>
                        <th className="w-36">Chi bộ sinh hoạt</th>
                        <th className="w-28 text-center">Ngày vào Đảng</th>
                        <th className="w-28 text-center">Ngày chính thức</th>
                        <th>Chức vụ Đảng / Đoàn thể</th>
                        <th className="w-28 text-center">Trạng thái</th>
                      </tr>
                    </thead>
                    <tbody>
                      {members
                        .filter(m => filterChiBo === 'Tất cả' || m.chiBo === filterChiBo)
                        .filter(m => !searchTerm || m.name.toLowerCase().includes(searchTerm.toLowerCase()) || m.soThe.includes(searchTerm))
                        .map((m, idx) => (
                          <tr key={m.id}>
                            <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                            <td className="font-bold text-xs">{m.name}</td>
                            <td className="text-center font-mono text-xs">{m.soThe}</td>
                            <td className="text-xs">{m.chiBo}</td>
                            <td className="text-center text-xs">{m.ngayVao}</td>
                            <td className="text-center text-xs">{m.ngayChinhThuc}</td>
                            <td className="text-xs">{m.chucVu}</td>
                            <td className="text-center">
                              <span className="badge bg-emerald-50 text-emerald-700 border border-emerald-200 text-[10px]">
                                {m.trangThai}
                              </span>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

        </main>

        {/* MODAL CHỈNH SỬA CHI TIẾT DÀNH CHO QUẢN TRỊ VIÊN (ADMIN QUICK EDIT MODAL) */}
        {adminEditModal.isOpen && (
          <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-[var(--bg-card)] border border-[var(--border-color)] rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden animate-in fade-in zoom-in-95">
              <div className="p-4 bg-red-600 text-white flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Edit className="w-5 h-5" />
                  <h3 className="font-bold text-sm">
                    {adminEditModal.type === 'meeting' ? 'Chỉnh Sửa Lịch Họp' : 'Chỉnh Sửa Đánh Giá Xếp Loại'} - {adminEditModal.chiBoName}
                  </h3>
                </div>
                <button 
                  onClick={() => setAdminEditModal(prev => ({ ...prev, isOpen: false }))}
                  className="text-white/80 hover:text-white font-black text-base cursor-pointer"
                >
                  ✕
                </button>
              </div>

              <form onSubmit={handleSaveAdminEditModal} className="p-5 space-y-4">
                <div className="text-xs text-[var(--text-muted)] bg-slate-50 dark:bg-slate-900/50 p-3 rounded-lg border border-[var(--border-color)]">
                  Tháng áp dụng: <b>Tháng {selectedMonth}/{selectedYear}</b> • Đơn vị: <b>{adminEditModal.chiBoName}</b>
                </div>

                {/* Date Picker & Format */}
                <div className="space-y-1.5">
                  <label className="block text-xs font-bold text-[var(--text-main)]">
                    1. Chọn ngày từ Lịch (*)
                  </label>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <input
                      type="date"
                      value={adminEditModal.dateValue}
                      onChange={(e) => {
                        const val = e.target.value;
                        const formatted = adminEditModal.type === 'meeting' 
                          ? formatVietnameseDateWithDay(val) 
                          : formatVietnameseDateOnly(val);
                        setAdminEditModal(prev => ({
                          ...prev,
                          dateValue: val,
                          dateFormatted: formatted
                        }));
                      }}
                      className="p-2 border rounded-lg text-xs"
                    />
                    <input
                      type="text"
                      value={adminEditModal.dateFormatted}
                      onChange={(e) => setAdminEditModal(prev => ({ ...prev, dateFormatted: e.target.value }))}
                      placeholder="Hiển thị: 03/10/2026 (thứ Bảy)"
                      className="p-2 border rounded-lg text-xs font-bold text-blue-600 dark:text-blue-400"
                    />
                  </div>
                </div>

                {adminEditModal.type === 'meeting' ? (
                  <>
                    <div className="space-y-1.5">
                      <label className="block text-xs font-bold text-[var(--text-main)]">2. Địa điểm họp (*)</label>
                      <input
                        type="text"
                        required
                        value={adminEditModal.location}
                        onChange={(e) => setAdminEditModal(prev => ({ ...prev, location: e.target.value }))}
                        className="p-2 border rounded-lg text-xs w-full"
                      />
                    </div>

                    <div className="space-y-1.5">
                      <label className="block text-xs font-bold text-[var(--text-main)]">3. Trạng thái phê duyệt</label>
                      <select
                        value={adminEditModal.status}
                        onChange={(e) => setAdminEditModal(prev => ({ ...prev, status: e.target.value }))}
                        className="p-2 border rounded-lg text-xs w-full font-bold"
                      >
                        <option value="Đã thống nhất">Đã thống nhất</option>
                        <option value="Đã đăng ký (Chờ duyệt)">Đã đăng ký (Chờ duyệt)</option>
                        <option value="Chờ đăng ký">Chờ đăng ký</option>
                      </select>
                    </div>

                    <div className="space-y-1.5">
                      <label className="block text-xs font-bold text-[var(--text-main)]">4. Ghi chú</label>
                      <textarea
                        rows={2}
                        value={adminEditModal.note}
                        onChange={(e) => setAdminEditModal(prev => ({ ...prev, note: e.target.value }))}
                        className="p-2 border rounded-lg text-xs w-full"
                      />
                    </div>
                  </>
                ) : (
                  <>
                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1.5">
                        <label className="block text-xs font-bold text-[var(--text-main)]">2. Số đảng viên dự / Tổng số</label>
                        <input
                          type="text"
                          value={adminEditModal.attendance}
                          onChange={(e) => setAdminEditModal(prev => ({ ...prev, attendance: e.target.value }))}
                          placeholder="VD: 18/18"
                          className="p-2 border rounded-lg text-xs font-bold w-full"
                        />
                      </div>

                      <div className="space-y-1.5">
                        <label className="block text-xs font-bold text-[var(--text-main)]">3. Điểm ĐG (0-100)</label>
                        <input
                          type="number"
                          min={0}
                          max={100}
                          value={adminEditModal.score}
                          onChange={(e) => {
                            const val = e.target.value;
                            const score = Number(val);
                            let cls = 'Tốt';
                            if (score >= 90) cls = 'Tốt';
                            else if (score >= 70) cls = 'Khá';
                            else if (score >= 50) cls = 'Trung bình';
                            else cls = 'Kém';
                            setAdminEditModal(prev => ({
                              ...prev,
                              score: val === '' ? '' : score,
                              classification: cls
                            }));
                          }}
                          className="p-2 border rounded-lg text-xs font-black text-center text-blue-700 w-full"
                        />
                      </div>
                    </div>

                    <div className="space-y-1.5">
                      <label className="block text-xs font-bold text-[var(--text-main)]">4. Mức xếp loại</label>
                      <select
                        value={adminEditModal.classification}
                        onChange={(e) => setAdminEditModal(prev => ({ ...prev, classification: e.target.value }))}
                        className="p-2 border rounded-lg text-xs w-full font-bold text-emerald-700"
                      >
                        <option value="Tốt">Tốt (Từ 90-100 điểm)</option>
                        <option value="Khá">Khá (Từ 70-89 điểm)</option>
                        <option value="Trung bình">Trung bình (Từ 50-69 điểm)</option>
                        <option value="Kém">Kém (Dưới 50 điểm)</option>
                      </select>
                    </div>

                    <div className="space-y-1.5">
                      <label className="block text-xs font-bold text-[var(--text-main)]">5. Thuyết minh điểm trừ / Ghi chú</label>
                      <textarea
                        rows={2}
                        value={adminEditModal.note}
                        onChange={(e) => setAdminEditModal(prev => ({ ...prev, note: e.target.value }))}
                        className="p-2 border rounded-lg text-xs w-full"
                      />
                    </div>
                  </>
                )}

                <div className="flex items-center justify-end gap-2 pt-2">
                  <button
                    type="button"
                    onClick={() => setAdminEditModal(prev => ({ ...prev, isOpen: false }))}
                    className="px-4 py-2 rounded-lg border border-[var(--border-color)] text-xs font-semibold cursor-pointer"
                  >
                    Đóng
                  </button>
                  <button
                    type="submit"
                    className="px-4 py-2 rounded-lg bg-red-600 hover:bg-red-700 text-white text-xs font-bold shadow flex items-center gap-1.5 cursor-pointer"
                  >
                    <Save className="w-4 h-4" />
                    Lưu & Đồng Bộ Lên Google Sheets
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* MODAL CÀI ĐẶT GOOGLE SHEETS API */}
        {isApiModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-[var(--bg-card)] border border-[var(--border-color)] rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden animate-in fade-in zoom-in-95">
              <div className="p-4 bg-emerald-600 text-white flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Database className="w-5 h-5" />
                  <h3 className="font-bold text-sm">Cài Đặt Kết Nối Google Sheets Online</h3>
                </div>
                <button 
                  onClick={() => setIsApiModalOpen(false)}
                  className="text-white/80 hover:text-white font-black text-base cursor-pointer"
                >
                  ✕
                </button>
              </div>

              <div className="p-5 space-y-4">
                <div>
                  <label className="block text-xs font-bold text-[var(--text-main)] mb-1">
                    Đường dẫn Google Apps Script Web App URL:
                  </label>
                  <input
                    type="url"
                    value={apiInputUrl}
                    onChange={(e) => setApiInputUrl(e.target.value)}
                    placeholder="https://script.google.com/macros/s/.../exec"
                    className="w-full text-xs font-mono p-2.5"
                  />
                </div>

                <div className="p-3 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200 text-xs text-emerald-800 dark:text-emerald-200 space-y-1.5">
                  <p className="font-bold">✨ Trạng thái kết nối hiện tại:</p>
                  <p className="text-[11px] font-mono break-all text-emerald-700 dark:text-emerald-300">
                    {googleScriptUrl || 'Chưa thiết lập URL mặc định'}
                  </p>
                  <p className="text-[11px] text-[var(--text-muted)]">
                    Lần đồng bộ gần nhất: <b>{lastSyncTime || 'Chưa thực hiện'}</b>
                  </p>
                </div>

                <div className="flex items-center justify-end gap-2 pt-2">
                  <button
                    onClick={() => {
                      setApiInputUrl(DEFAULT_GOOGLE_SCRIPT_URL);
                      setGoogleScriptUrl(DEFAULT_GOOGLE_SCRIPT_URL);
                      localStorage.setItem('qltt_google_script_url', DEFAULT_GOOGLE_SCRIPT_URL);
                      fetchFromGoogleSheets(DEFAULT_GOOGLE_SCRIPT_URL);
                      alert('Đã khôi phục URL Google Apps Script mặc định thành công!');
                    }}
                    className="px-3 py-2 rounded-lg border border-[var(--border-color)] text-xs font-semibold cursor-pointer"
                  >
                    Khôi phục mặc định
                  </button>

                  <button
                    onClick={() => {
                      setGoogleScriptUrl(apiInputUrl);
                      localStorage.setItem('qltt_google_script_url', apiInputUrl);
                      setIsApiModalOpen(false);
                      if (apiInputUrl) {
                        fetchFromGoogleSheets(apiInputUrl);
                      }
                      alert('Đã lưu cấu hình kết nối Google Sheets!');
                    }}
                    className="px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shadow cursor-pointer"
                  >
                    Lưu cấu hình & Đồng bộ
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* MODAL SHARE LINK CHO CÁC CHI BỘ */}
        {isShareModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
            <div className="bg-[var(--bg-card)] border border-[var(--border-color)] rounded-2xl w-full max-w-lg shadow-2xl overflow-hidden animate-in fade-in zoom-in-95">
              <div className="p-4 bg-indigo-600 text-white flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <Share2 className="w-5 h-5" />
                  <h3 className="font-bold text-sm">Chia Sẻ Link Cho Các Chi Bộ (Tháng {selectedMonth}/{selectedYear})</h3>
                </div>
                <button 
                  onClick={() => setIsShareModalOpen(false)}
                  className="text-white/80 hover:text-white font-black text-base cursor-pointer"
                >
                  ✕
                </button>
              </div>

              {/* TABS TRONG MODAL SHARE */}
              <div className="flex border-b border-[var(--border-color)] bg-slate-50 dark:bg-slate-900/50 p-2 gap-2">
                <button
                  onClick={() => setShareTab('register')}
                  className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                    shareTab === 'register' 
                      ? 'bg-red-600 text-white shadow-sm' 
                      : 'text-[var(--text-muted)] hover:bg-slate-200 dark:hover:bg-slate-800'
                  }`}
                >
                  📝 1. Link Đăng Ký Lịch Họp (Trước ngày 20)
                </button>
                <button
                  onClick={() => setShareTab('report')}
                  className={`flex-1 py-2 text-xs font-bold rounded-lg transition-all cursor-pointer ${
                    shareTab === 'report' 
                      ? 'bg-blue-600 text-white shadow-sm' 
                      : 'text-[var(--text-muted)] hover:bg-slate-200 dark:hover:bg-slate-800'
                  }`}
                >
                  ⭐ 2. Link Báo Cáo ĐGXL (Sau khi họp)
                </button>
              </div>

              <div className="p-5 space-y-4">
                {shareTab === 'register' ? (
                  <>
                    <div>
                      <label className="block text-xs font-bold text-[var(--text-main)] mb-1">
                        1. Đường link gửi Zalo cho các Đội đăng ký ngày họp Tháng {selectedMonth}/{selectedYear}:
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          readOnly
                          value={currentRegisterLink}
                          className="w-full text-xs font-mono bg-slate-50 dark:bg-slate-900 p-2"
                        />
                        <button
                          onClick={() => copyShareLink(currentRegisterLink)}
                          className="px-3 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-xs font-bold flex items-center gap-1 cursor-pointer shrink-0"
                        >
                          {copiedLink ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                          {copiedLink ? 'Đã chép' : 'Sao chép'}
                        </button>
                      </div>
                    </div>

                    <div className="border border-[var(--border-color)] rounded-xl p-3 bg-slate-50 dark:bg-slate-900/50">
                      <div className="text-xs font-bold text-[var(--text-main)] mb-1">2. Mẫu tin nhắn soạn sẵn gửi Zalo nhóm Đảng bộ:</div>
                      <div className="text-xs text-[var(--text-muted)] p-2 bg-white dark:bg-slate-800 rounded border border-[var(--border-color)] font-sans leading-relaxed">
                        Kính gửi các đồng chí Bí thư Chi bộ trực thuộc Đảng bộ QLTT,<br/><br/>
                        Đề nghị các Chi bộ thực hiện đăng ký lịch họp sinh hoạt chi bộ <b>Tháng {selectedMonth}/{selectedYear}</b> trước ngày 20 theo đường link dưới đây:<br/>
                        👉 <span className="text-blue-600 underline font-mono text-[11px]">{currentRegisterLink}</span><br/><br/>
                        Trân trọng cảm ơn!
                      </div>
                      <button
                        onClick={() => copyZaloMessage('register')}
                        className="mt-2 w-full py-2 bg-red-600 hover:bg-red-700 text-white font-bold text-xs rounded-md flex items-center justify-center gap-1 cursor-pointer"
                      >
                        {copiedMessage ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                        {copiedMessage ? 'Đã sao chép tin nhắn Zalo' : 'Sao chép Toàn bộ Tin Nhắn gửi Zalo'}
                      </button>
                    </div>
                  </>
                ) : (
                  <>
                    <div>
                      <label className="block text-xs font-bold text-[var(--text-main)] mb-1">
                        1. Đường link gửi Zalo cho các Đội báo cáo kết quả ĐGXL Tháng {selectedMonth}/{selectedYear}:
                      </label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          readOnly
                          value={currentReportLink}
                          className="w-full text-xs font-mono bg-slate-50 dark:bg-slate-900 p-2"
                        />
                        <button
                          onClick={() => copyShareLink(currentReportLink)}
                          className="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold flex items-center gap-1 cursor-pointer shrink-0"
                        >
                          {copiedLink ? <Check className="w-3.5 h-3.5" /> : <Copy className="w-3.5 h-3.5" />}
                          {copiedLink ? 'Đã chép' : 'Sao chép'}
                        </button>
                      </div>
                    </div>

                    <div className="border border-[var(--border-color)] rounded-xl p-3 bg-slate-50 dark:bg-slate-900/50">
                      <div className="text-xs font-bold text-[var(--text-main)] mb-1">2. Mẫu tin nhắn soạn sẵn gửi Zalo nhóm Đảng bộ:</div>
                      <div className="text-xs text-[var(--text-muted)] p-2 bg-white dark:bg-slate-800 rounded border border-[var(--border-color)] font-sans leading-relaxed">
                        Kính gửi các đồng chí Bí thư Chi bộ trực thuộc Đảng bộ QLTT,<br/><br/>
                        Sau khi hoàn thành buổi sinh hoạt chi bộ <b>Tháng {selectedMonth}/{selectedYear}</b>, đề nghị các Chi bộ gửi Báo cáo Đánh giá, xếp loại sinh hoạt chi bộ (theo Hướng dẫn 01-HD/TU) tại đường link dưới đây để Đảng ủy tổng hợp:<br/>
                        👉 <span className="text-blue-600 underline font-mono text-[11px]">{currentReportLink}</span><br/><br/>
                        Trân trọng cảm ơn!
                      </div>
                      <button
                        onClick={() => copyZaloMessage('report')}
                        className="mt-2 w-full py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-md flex items-center justify-center gap-1 cursor-pointer"
                      >
                        {copiedMessage ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                        {copiedMessage ? 'Đã sao chép tin nhắn Zalo' : 'Sao chép Toàn bộ Tin Nhắn gửi Zalo'}
                      </button>
                    </div>
                  </>
                )}

                <div className="p-3 rounded-lg bg-blue-50 dark:bg-blue-950/40 text-xs text-blue-800 dark:text-blue-200">
                  💡 <b>Tiện ích:</b> Link trên đã được gắn sẵn Tháng {selectedMonth}/{selectedYear}. Chi bộ bấm vào link sẽ mở đúng Tháng {selectedMonth} để đăng ký hoặc báo cáo, đồng thời có thể tự đổi sang bất kỳ tháng nào tùy ý!
                </div>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
