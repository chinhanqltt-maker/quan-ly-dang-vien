import React, { useState, useEffect, useMemo } from 'react';
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
  Eye
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

export default function App() {
  const [darkMode, setDarkMode] = useState(false);

  // Check URL params for Member Registration / Report Mode
  const urlParams = new URLSearchParams(window.location.search);
  const isPublicRegisterMode = urlParams.get('mode') === 'register';
  const isPublicReportMode = urlParams.get('mode') === 'report';
  const paramChiBoId = urlParams.get('chibo');

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
  const [selectedYear, setSelectedYear] = useState(2026);
  const [selectedMonth, setSelectedMonth] = useState(9);

  const [searchTerm, setSearchTerm] = useState('');
  const [filterChiBo, setFilterChiBo] = useState('Tất cả');

  // Modal Share Link
  const [isShareModalOpen, setIsShareModalOpen] = useState(false);
  const [shareTab, setShareTab] = useState('register'); // 'register' | 'report'
  const [copiedLink, setCopiedLink] = useState(false);
  const [copiedMessage, setCopiedMessage] = useState(false);

  // Print Mode State ('schedule' | 'dgxl')
  const [printDocType, setPrintDocType] = useState('schedule');

  // Form tự đăng ký lịch họp của Đội (mode=register)
  const initialSelectedBranchId = paramChiBoId ? Number(paramChiBoId) : 1;
  const initialBranchObj = branchDetails.find(b => b.id === initialSelectedBranchId) || branchDetails[0];

  const [teamSelectChiBo, setTeamSelectChiBo] = useState(initialSelectedBranchId);
  const [teamDate, setTeamDate] = useState('');
  const [teamDayOfWeek, setTeamDayOfWeek] = useState('thứ Hai');
  const [teamLocation, setTeamLocation] = useState(initialBranchObj.diaDiem);
  const [teamNote, setTeamNote] = useState('');
  const [teamSubmitted, setTeamSubmitted] = useState(false);

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
  };

  const handleApproveAndArchive = () => {
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

    alert(`Đã duyệt thống nhất với Lãnh đạo và tự động lưu vào 'Kho Lưu Trữ Thông Báo'! Đồng chí có thể bấm 'Xuất File Word (.DOC)' hoặc 'In Trình Ký'.`);
  };

  // Dữ liệu ĐGXL chuẩn 13 chi bộ (Chi bộ 1 -> Chi bộ 12, Chi bộ KP)
  const currentMonthDGXL = useMemo(() => {
    if (dgxlData[currentKey]) {
      return dgxlData[currentKey];
    }
    // Khởi tạo theo danh sách mẫu: các chi bộ 1, 3, 7 có mẫu như hình, các chi bộ còn lại để chờ nhập
    return STANDARD_13_BRANCHES.map(b => {
      const schedule = currentMonthSchedules.find(s => s.chiBoId === b.id) || {};
      let meetingDate = schedule.thoiGian || '';
      if (meetingDate && meetingDate.includes('(')) {
        meetingDate = meetingDate.split('(')[0].trim();
      }
      
      // Khởi tạo mặc định một số chi bộ mẫu hoặc dữ liệu đã có
      if (b.id === 1) {
        return {
          chiBoId: b.id,
          chiBo: b.shortName,
          sl: `${b.sl}/${b.sl}`,
          ngayHop: meetingDate || '03/6/2026',
          diemDG: 100,
          mucXepLoai: 'Tốt',
          ghiChuTruDiem: ''
        };
      }
      if (b.id === 3) {
        return {
          chiBoId: b.id,
          chiBo: b.shortName,
          sl: `${b.sl}/${b.sl}`,
          ngayHop: meetingDate || '3/9/2026',
          diemDG: 100,
          mucXepLoai: 'Tốt',
          ghiChuTruDiem: ''
        };
      }
      if (b.id === 7) {
        return {
          chiBoId: b.id,
          chiBo: b.shortName,
          sl: `${b.sl}/${b.sl}`,
          ngayHop: meetingDate || '4/9/2026',
          diemDG: 100,
          mucXepLoai: 'Tốt',
          ghiChuTruDiem: ''
        };
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

  // Submit từ form đăng ký lịch họp (mode=register)
  const handleTeamSubmitRegistration = (e) => {
    e.preventDefault();
    if (!teamDate) {
      alert('Vui lòng chọn ngày họp!');
      return;
    }
    const [yyyy, mm, dd] = teamDate.split('-');
    const timeFormatted = `${dd}/${mm}/${yyyy} (${teamDayOfWeek})`;
    
    const targetBranch = branchDetails.find(b => b.id === teamSelectChiBo);
    const loc = teamLocation || (targetBranch ? targetBranch.diaDiem : 'Tại đơn vị');

    handleUpdateSchedule(teamSelectChiBo, 'thoiGian', timeFormatted);
    handleUpdateSchedule(teamSelectChiBo, 'diaDiem', loc);
    if (teamNote) {
      handleUpdateSchedule(teamSelectChiBo, 'ghiChu', teamNote);
    }

    setTeamSubmitted(true);
    setTimeout(() => setTeamSubmitted(false), 5000);
  };

  // Submit từ form báo cáo ĐGXL sau họp (mode=report)
  const handleTeamSubmitReport = (e) => {
    e.preventDefault();
    let dateStr = reportMeetingDate;
    if (reportMeetingDate && reportMeetingDate.includes('-')) {
      const [yyyy, mm, dd] = reportMeetingDate.split('-');
      dateStr = `${Number(dd)}/${Number(mm)}/${yyyy}`;
    }

    handleUpdateDGXL(reportBranchId, 'sl', reportAttendance);
    if (dateStr) {
      handleUpdateDGXL(reportBranchId, 'ngayHop', dateStr);
    }
    handleUpdateDGXL(reportBranchId, 'diemDG', Number(reportScore));
    handleUpdateDGXL(reportBranchId, 'mucXepLoai', reportClassification);
    if (reportDeductionNote) {
      handleUpdateDGXL(reportBranchId, 'ghiChuTruDiem', reportDeductionNote);
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

          <!-- Dòng Tổng số đảng viên và thuyết minh biến động -->
          <p style="font-size: 11pt; margin-top: 3pt; margin-bottom: 6pt; font-style: italic;">
            ${currentMonthNote}
          </p>

          <!-- Footer Table -->
          <table style="width: 100%; border: none; margin-top: 4pt;">
            <tr>
              <td style="width: 50%; vertical-align: top; border: none; line-height: 100%;">
                <p style="font-size: 12pt; font-weight: bold; margin: 0; line-height: 100%;"><u>Nơi nhận:</u></p>
                <p style="font-size: 12pt; line-height: 100%; margin: 0;">
                  - Đảng ủy Sở;<br/>
                  - Đảng ủy Chi cục QLTT,<br/>
                  - Lưu: ĐU.
                </p>
              </td>
              <td style="width: 50%; text-align: center; vertical-align: top; border: none; line-height: 100%;">
                <p style="font-size: 12pt; font-weight: bold; margin: 0; line-height: 100%;">T/M ĐẢNG ỦY</p>
                <p style="font-size: 12pt; font-weight: bold; margin: 0; line-height: 100%;">BÍ THƯ</p>
                <div style="height: 38pt;"></div>
                <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">Nguyễn Trung Tiến</p>
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    saveAs(blob, `${selectedMonth}. BC_T${selectedMonth}_xếp loại chi bộ theo HD 01 TU.doc`);
  };

  const handleExportSixMonthsToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(sixMonthsData.map((d, idx) => ({
      'STT': idx + 1,
      'Tên chi bộ': d.chiBo,
      'Đảng số': d.dangSo,
      'Số lượng ĐV dự sinh hoạt lệ': d.soLgDuHop,
      'Phát biểu ý kiến': d.phatBieuYkien,
      'Dân vận khéo': d.danVanKheo,
      'Học tập Chuyên đề': d.hocTapChuyenDe,
      'Đọc bài viết': d.docBaiViet,
      'KH Sinh hoạt CĐ': d.khSinhHoatCD,
      'KNĐ': d.knd,
      'KT GS': d.ktgs,
      'Ghi chú': d.ghiChu,
      'Khác': d.khac
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `TongHop_06Thang_${selectedYear}`);
    XLSX.writeFile(wb, `Bao_Cao_Tong_Hop_06_Thang_${selectedYear}.xlsx`);
  };

  const totalMembersCount = currentMonthSchedules.reduce((sum, i) => sum + (Number(i.sl) || 0), 0);

  // XUẤT FILE WORD .DOC LỊCH HỌP CHUẨN ĐỊNH DẠNG MICROSOFT WORD 100%
  const handleExportWordDoc = () => {
    const tableRows = currentMonthSchedules.map((item, idx) => {
      const isThuyOrThuy = item.biThu && (item.biThu.includes('Trần Thị Thu Thanh Thủy') || item.biThu.includes('Nguyễn Phúc Xuân Thụy'));
      const isThuyLong = item.biThu && item.biThu.includes('Trần Thị Thu Thanh Thủy');
      const biThuStyle = isThuyOrThuy 
        ? `font-size: 12pt; ${isThuyLong ? 'letter-spacing: -0.3pt;' : ''} white-space: nowrap;` 
        : 'font-size: 13pt;';
      
      let datePart = item.thoiGian || '';
      let dowPart = '';
      if (item.thoiGian && item.thoiGian.includes('(')) {
        const parts = item.thoiGian.split('(');
        datePart = parts[0].trim();
        dowPart = '(' + parts[1].trim();
      }

      return `
      <tr style="height: 22px;">
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${idx + 1}</td>
        <td style="border: 1px solid black; padding: 2px 4px; font-size: 11.5pt;">${item.chiBo}</td>
        <td style="border: 1px solid black; padding: 2px; text-align: center; font-size: 11pt;">${item.sl}</td>
        <td style="border: 1px solid black; padding: 2px 3px; text-align: center; line-height: 100%;">
          <span style="font-size: 11.5pt;">${datePart}</span><br/><span style="font-size: 10.5pt;">${dowPart}</span>
        </td>
        <td style="border: 1px solid black; padding: 2px 4px; text-align: center; font-size: 11.5pt;">${item.diaDiem}</td>
        <td style="border: 1px solid black; padding: 1px 2px; text-align: center; line-height: 100%;">
          <b style="${biThuStyle}">${item.biThu}</b><br/><span style="font-size: 10.5pt;">${item.sdt || ''}</span>
        </td>
      </tr>
      `;
    }).join('');

    const wordHtml = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>THÔNG BÁO LỊCH HỌP THÁNG ${selectedMonth}-${selectedYear}</title>
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
            margin: 1.0cm 1.5cm 1.0cm 2.0cm;
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
          <table style="width: 100%; border: none; margin-bottom: 2pt;">
            <tr>
              <td style="width: 56%; text-align: center; vertical-align: top; border: none; font-size: 11pt; white-space: nowrap; line-height: 100%;">
                ĐẢNG BỘ SỞ CÔNG THƯƠNG TỈNH AN GIANG<br/>
                <b>ĐẢNG ỦY BỘ PHẬN CHI CỤC QUẢN LÝ THỊ TRƯỜNG</b><br/>
                *<br/>
                <span style="font-size: 12pt;">Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; -TB/ĐU</span>
              </td>
              <td style="width: 44%; text-align: center; vertical-align: top; border: none; font-size: 12pt; white-space: nowrap; line-height: 100%;">
                <b>ĐẢNG CỘNG SẢN VIỆT NAM</b><br/><br/>
                <i>An Giang, ngày &nbsp;&nbsp;&nbsp;&nbsp; tháng ${String(selectedMonth === 1 ? 12 : selectedMonth - 1).padStart(2, '0')} năm ${selectedMonth === 1 ? selectedYear - 1 : selectedYear}</i>
              </td>
            </tr>
          </table>

          <!-- Title -->
          <div style="text-align: center; margin-top: 2pt; margin-bottom: 2pt;">
            <p style="font-size: 14pt; font-weight: bold; margin: 0; line-height: 100%;">THÔNG BÁO</p>
            <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">Lịch sinh hoạt lệ tháng ${String(selectedMonth).padStart(2, '0')}/${selectedYear} của các chi bộ trực thuộc</p>
            <p style="font-size: 11pt; margin: 0; line-height: 100%;">-----</p>
          </div>

          <!-- Paragraphs Cỡ chữ 14 chuẩn quy định Spacing 0 Single -->
          <p style="font-size: 14pt; text-indent: 30pt; text-align: justify; margin-top: 2pt; margin-bottom: 2pt; line-height: 100%;">
            Căn cứ Quy chế làm việc của Đảng ủy Chi cục Quản lý thị trường và Quy chế làm việc của các chi bộ trực thuộc nhiệm kỳ 2025-2030.
          </p>
          <p style="font-size: 14pt; text-indent: 30pt; text-align: justify; margin-top: 2pt; margin-bottom: 4pt; line-height: 100%;">
            Theo đăng ký lịch sinh hoạt lệ chi bộ tháng ${String(selectedMonth).padStart(2, '0')}/${selectedYear}. Đảng ủy bộ phận Chi cục Quản lý thị trường thông báo thời gian, địa điểm sinh hoạt của các chi bộ, như sau:
          </p>

          <!-- 13 Chi bộ Table -->
          <table style="width: 100%; border: 1px solid black; font-size: 10.5pt; margin-top: 2pt; margin-bottom: 4pt;">
            <tr style="text-align: center; font-weight: bold; font-size: 12pt; background-color: #f2f2f2; height: 20px;">
              <td style="border: 1px solid black; padding: 2px; width: 6%;">S<br/>TT</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 22%;">Chi bộ</td>
              <td style="border: 1px solid black; padding: 2px; width: 8%;">Số<br/>lượng</td>
              <td style="border: 1px solid black; padding: 2px 3px; width: 22%;">Thời<br/>gian</td>
              <td style="border: 1px solid black; padding: 2px 4px; width: 18%;">Địa điểm</td>
              <td style="border: 1px solid black; padding: 2px 3px; width: 24%;">Bí thư,<br/>điện thoại</td>
            </tr>
            ${tableRows}
            <tr style="font-weight: bold; text-align: center; height: 20px;">
              <td colspan="2" style="border: 1px solid black; padding: 2px 4px; text-align: center;">TỔNG SỐ</td>
              <td style="border: 1px solid black; padding: 2px;">${totalMembersCount}</td>
              <td style="border: 1px solid black; padding: 2px;"></td>
              <td style="border: 1px solid black; padding: 2px;"></td>
              <td style="border: 1px solid black; padding: 2px;"></td>
            </tr>
          </table>

          <!-- Footer Table -->
          <table style="width: 100%; border: none; margin-top: 2pt;">
            <tr>
              <td style="width: 50%; vertical-align: top; border: none; line-height: 100%;">
                <p style="font-size: 12pt; font-weight: bold; margin: 0; line-height: 100%;"><u>Nơi nhận:</u></p>
                <p style="font-size: 12pt; line-height: 100%; margin: 0;">
                  - Đảng ủy Sở Công Thương;<br/>
                  - Bí thư các Chi bộ trực thuộc;<br/>
                  - Lưu: Đảng ủy.
                </p>
              </td>
              <td style="width: 50%; text-align: center; vertical-align: top; border: none; line-height: 100%;">
                <p style="font-size: 12pt; font-weight: bold; margin: 0; line-height: 100%;">T/M ĐẢNG ỦY</p>
                <p style="font-size: 12pt; font-weight: bold; margin: 0; line-height: 100%;">BÍ THƯ</p>
                <div style="height: 36pt;"></div>
                <p style="font-size: 13pt; font-weight: bold; margin: 0; line-height: 100%;">Nguyễn Trung Tiến</p>
              </td>
            </tr>
          </table>
        </div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + wordHtml], { type: 'application/msword' });
    saveAs(blob, `${selectedMonth}. TB_họp chi bộ T${selectedMonth}-${selectedYear}_13 chi bộ.doc`);
  };

  const handlePrintSchedule = () => {
    setPrintDocType('schedule');
    const originalTitle = document.title;
    document.title = "";
    setTimeout(() => {
      window.print();
      document.title = originalTitle;
    }, 100);
  };

  const handlePrintDGXL = () => {
    setPrintDocType('dgxl');
    const originalTitle = document.title;
    document.title = "";
    setTimeout(() => {
      window.print();
      document.title = originalTitle;
    }, 100);
  };

  const sampleZaloRegisterMessage = `[THÔNG BÁO ĐẢNG ỦY BỘ PHẬN CHI CỤC QLTT]
Kính gửi: Bí thư các Chi bộ trực thuộc (Đội 1 đến Đội 12 và Khối phòng).
Thực hiện Quy chế làm việc, đề nghị các Chi bộ chủ động đăng ký lịch sinh hoạt lệ tháng ${selectedMonth}/${selectedYear} trước ngày 20 để Đảng ủy tổng hợp xin ý kiến Lãnh đạo và ban hành Thông báo chính thức.
👉 Link đăng ký trực tuyến: ${window.location.origin + window.location.pathname}?mode=register`;

  const sampleZaloReportMessage = `[BÁO CÁO ĐÁNH GIÁ XẾP LOẠI CHI BỘ THÁNG ${selectedMonth}/${selectedYear}]
Kính gửi: Bí thư các Chi bộ trực thuộc (Đội 1 đến Đội 12 và Khối phòng).
Theo Hướng dẫn 01-HD/TU, đề nghị các Chi bộ sau khi tổ chức buổi sinh hoạt lệ hoàn thành tự chấm điểm và gửi Báo cáo Đánh giá Xếp loại về Đảng ủy bộ phận.
👉 Link gửi Báo cáo ĐGXL trực tuyến: ${window.location.origin + window.location.pathname}?mode=report`;

  // GIAO DIỆN DÀNH CHO CÁC ĐỘI TỰ ĐĂNG KÝ LỊCH HỌP (mode=register)
  if (isPublicRegisterMode) {
    const selectedBranchInfo = branchDetails.find(b => b.id === teamSelectChiBo) || branchDetails[0];

    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4 font-sans">
        <div className="bg-white border border-slate-200 rounded-2xl w-full max-w-xl shadow-xl overflow-hidden">
          <div className="bg-red-600 p-6 text-white text-center relative">
            <div className="w-12 h-12 mx-auto rounded-xl bg-yellow-400 text-red-700 flex items-center justify-center font-black text-2xl shadow mb-2">
              ★
            </div>
            <h2 className="text-base font-bold uppercase tracking-wide">ĐẢNG ỦY BỘ PHẬN CHI CỤC QLTT</h2>
            <h3 className="text-lg font-black mt-1">ĐĂNG KÝ LỊCH SINH HOẠT LỆ CHI BỘ</h3>
            <p className="text-xs text-red-100 mt-1">Tháng {selectedMonth} năm {selectedYear}</p>
          </div>

          <form onSubmit={handleTeamSubmitRegistration} className="p-6 space-y-4">
            {teamSubmitted ? (
              <div className="p-5 rounded-xl bg-emerald-50 border border-emerald-200 text-center space-y-2">
                <CheckCircle2 className="w-12 h-12 text-emerald-600 mx-auto" />
                <h4 className="font-bold text-emerald-800 text-base">Đăng ký thành công!</h4>
                <p className="text-xs text-emerald-700">
                  Lịch sinh hoạt của <b>{selectedBranchInfo.name}</b> tại <b>{teamLocation}</b> đã được lưu vào hệ thống Đảng ủy bộ phận.
                </p>
                <button
                  type="button"
                  onClick={() => setTeamSubmitted(false)}
                  className="mt-2 text-xs text-blue-600 font-bold underline cursor-pointer"
                >
                  Đăng ký lại hoặc chọn Chi bộ khác
                </button>
              </div>
            ) : (
              <>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">1. Chọn Chi bộ trực thuộc (*)</label>
                  <select
                    value={teamSelectChiBo}
                    onChange={(e) => setTeamSelectChiBo(Number(e.target.value))}
                    className="w-full p-2.5 border rounded-lg text-sm font-bold text-red-700 bg-red-50/50 border-red-200"
                  >
                    {branchDetails.map(b => (
                      <option key={b.id} value={b.id}>{b.name} (Bí thư: {b.biThu} - {b.sl} ĐV)</option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">2. Chọn Ngày họp (*)</label>
                    <input
                      type="date"
                      required
                      min={`${selectedYear}-${String(selectedMonth).padStart(2, '0')}-01`}
                      max={`${selectedYear}-${String(selectedMonth).padStart(2, '0')}-${new Date(selectedYear, selectedMonth, 0).getDate()}`}
                      value={teamDate}
                      onChange={(e) => {
                        setTeamDate(e.target.value);
                        if (e.target.value) {
                          const d = new Date(e.target.value);
                          const days = ['Chủ Nhật', 'thứ Hai', 'thứ Ba', 'thứ Tư', 'thứ Năm', 'thứ Sáu', 'thứ Bảy'];
                          setTeamDayOfWeek(days[d.getDay()]);
                        }
                      }}
                      className="w-full p-2.5 border rounded-lg text-sm bg-white"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">Thứ trong tuần</label>
                    <input
                      type="text"
                      readOnly
                      value={teamDayOfWeek}
                      className="w-full p-2.5 border rounded-lg text-sm bg-slate-100 font-semibold"
                    />
                  </div>
                </div>

                <div>
                  <div className="flex items-center justify-between mb-1">
                    <label className="text-xs font-bold text-slate-700 flex items-center gap-1">
                      <MapPin className="w-3.5 h-3.5 text-red-600" />
                      3. Địa điểm tổ chức sinh hoạt
                    </label>
                    <span className="text-[11px] text-emerald-600 font-semibold">(Mặc định theo dữ liệu của Đội)</span>
                  </div>
                  <input
                    type="text"
                    required
                    value={teamLocation}
                    onChange={(e) => setTeamLocation(e.target.value)}
                    className="w-full p-2.5 border rounded-lg text-sm bg-white font-medium"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    4. Ghi chú biến động đảng viên (nếu có)
                  </label>
                  <input
                    type="text"
                    value={teamNote}
                    onChange={(e) => setTeamNote(e.target.value)}
                    placeholder="VD: Có 01 đ/c miễn sinh hoạt hoặc chuyển công tác tạm thời..."
                    className="w-full p-2.5 border rounded-lg text-xs bg-white"
                  />
                </div>

                <div className="p-3 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-800">
                  📌 <b>Lưu ý:</b> Đề nghị các Chi bộ hoàn thành đăng ký trước ngày 20 hàng tháng.
                </div>

                <button
                  type="submit"
                  className="w-full py-3 rounded-xl bg-red-600 hover:bg-red-700 text-white font-bold text-sm shadow-md flex items-center justify-center gap-2 cursor-pointer transition-all"
                >
                  <Send className="w-4 h-4" />
                  Gửi Đăng Ký Lịch Họp
                </button>
              </>
            )}

            <div className="flex justify-between items-center pt-2 text-xs">
              <a 
                href={window.location.pathname + '?mode=report'}
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
            <p className="text-xs text-blue-100 mt-1">Theo Hướng dẫn 01-HD/TU • Tháng {selectedMonth} năm {selectedYear}</p>
          </div>

          <form onSubmit={handleTeamSubmitReport} className="p-6 space-y-4">
            {reportSubmitted ? (
              <div className="p-5 rounded-xl bg-emerald-50 border border-emerald-200 text-center space-y-3">
                <CheckCircle2 className="w-12 h-12 text-emerald-600 mx-auto" />
                <h4 className="font-bold text-emerald-800 text-base">Gửi báo cáo thành công!</h4>
                <p className="text-xs text-emerald-700 leading-relaxed">
                  Kết quả đánh giá xếp loại sinh hoạt của <b>{selectedBranchInfo.name}</b> (Điểm: <b>{reportScore}</b> - Xếp loại: <b>{reportClassification}</b>) đã được lưu vào hệ thống Đảng ủy.
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
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">1. Chọn Chi bộ báo cáo (*)</label>
                  <select
                    value={reportBranchId}
                    onChange={(e) => setReportBranchId(Number(e.target.value))}
                    className="w-full p-2.5 border rounded-lg text-sm font-bold text-blue-800 bg-blue-50/50 border-blue-200"
                  >
                    {branchDetails.map(b => (
                      <option key={b.id} value={b.id}>{b.name} (Bí thư: {b.biThu} - {b.sl} ĐV)</option>
                    ))}
                  </select>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">2. Số đảng viên dự / Tổng số (*)</label>
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
                    <label className="block text-xs font-bold text-slate-700 mb-1">3. Ngày họp thực tế (*)</label>
                    <input
                      type="date"
                      required
                      value={reportMeetingDate}
                      onChange={(e) => setReportMeetingDate(e.target.value)}
                      className="w-full p-2.5 border rounded-lg text-sm bg-white"
                    />
                  </div>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 mb-1">4. Điểm tự đánh giá (0-100) (*)</label>
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
                    <label className="block text-xs font-bold text-slate-700 mb-1">5. Mức xếp loại chi bộ (*)</label>
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
                    6. Thuyết minh điểm trừ / Ghi chú biến động (nếu có)
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
                  className="w-full py-3 rounded-xl bg-blue-700 hover:bg-blue-800 text-white font-bold text-sm shadow-md flex items-center justify-center gap-2 cursor-pointer transition-all"
                >
                  <Send className="w-4 h-4" />
                  Gửi Báo Cáo Đánh Giá Xếp Loại
                </button>
              </>
            )}

            <div className="flex justify-between items-center pt-2 text-xs">
              <a 
                href={window.location.pathname + '?mode=register'}
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

            <div className="flex items-center gap-2.5">
              {activeTab === 'dgxl' ? (
                <>
                  <button 
                    onClick={handleExportDGXLWordDoc}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-blue-700 hover:bg-blue-800 text-white shadow-sm transition-all cursor-pointer"
                    title="Tải về file Word (.DOC) ĐGXL chuẩn 1 trang A4 y hệt mẫu"
                  >
                    <FileText className="w-4 h-4" />
                    Xuất File Word (.DOC) ĐGXL
                  </button>

                  <button 
                    onClick={handlePrintDGXL}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-red-600 hover:bg-red-700 text-white shadow-sm transition-all cursor-pointer"
                    title="In bảng ĐGXL trình ký Bí thư Đảng ủy (Chuẩn 1 trang A4)"
                  >
                    <Printer className="w-4 h-4" />
                    In Báo Cáo ĐGXL (1 Trang)
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
                    Tải File Word (.DOC) Lịch Họp
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
          <div className="hidden print-page text-black bg-white" style={{ fontFamily: '"Times New Roman", Times, serif', width: '100%', fontSize: '13pt', lineHeight: '1.2' }}>
            <table style={{ width: '100%', border: 'none', borderCollapse: 'collapse', marginBottom: '4pt' }}>
              <tbody>
                <tr style={{ border: 'none' }}>
                  <td style={{ width: '56%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0, whiteSpace: 'nowrap' }}>
                    <div style={{ fontSize: '11pt' }}>ĐẢNG BỘ SỞ CÔNG THƯƠNG TỈNH AN GIANG</div>
                    <div style={{ fontSize: '11pt', fontWeight: 'bold' }}>ĐẢNG ỦY BỘ PHẬN CHI CỤC QUẢN LÝ THỊ TRƯỜNG</div>
                    <div style={{ fontSize: '10pt', margin: '1px 0' }}>*</div>
                    <div style={{ fontSize: '12pt' }}>Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; -TB/ĐU</div>
                  </td>
                  <td style={{ width: '44%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0, whiteSpace: 'nowrap' }}>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}>ĐẢNG CỘNG SẢN VIỆT NAM</div>
                    <div style={{ height: '12pt' }}></div>
                    <div style={{ fontSize: '12pt', fontStyle: 'italic' }}>
                      An Giang, ngày &nbsp;&nbsp;&nbsp;&nbsp; tháng {String(selectedMonth === 1 ? 12 : selectedMonth - 1).padStart(2, '0')} năm {selectedMonth === 1 ? selectedYear - 1 : selectedYear}
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>

            <div style={{ textAlign: 'center', margin: '4pt 0 2pt 0' }}>
              <div style={{ fontSize: '14pt', fontWeight: 'bold' }}>THÔNG BÁO</div>
              <div style={{ fontSize: '14pt', fontWeight: 'bold' }}>
                Lịch sinh hoạt lệ tháng {String(selectedMonth).padStart(2, '0')}/{selectedYear} của các chi bộ trực thuộc
              </div>
              <div style={{ fontSize: '12pt' }}>-----</div>
            </div>

            <div style={{ textIndent: '30pt', textAlign: 'justify', marginBottom: '2pt', fontSize: '14pt', lineHeight: '1.25' }}>
              Căn cứ Quy chế làm việc của Đảng ủy Chi cục Quản lý thị trường và Quy chế làm việc của các chi bộ trực thuộc nhiệm kỳ 2025-2030.
            </div>
            <div style={{ textIndent: '30pt', textAlign: 'justify', marginBottom: '4pt', fontSize: '14pt', lineHeight: '1.25' }}>
              Theo đăng ký lịch sinh hoạt lệ chi bộ tháng {String(selectedMonth).padStart(2, '0')}/{selectedYear}. Đảng ủy bộ phận Chi cục Quản lý thị trường thông báo thời gian, địa điểm sinh hoạt của các chi bộ, như sau:
            </div>

            <table className="schedule-table" style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid black', fontSize: '11pt', margin: '3pt 0' }}>
              <thead>
                <tr style={{ textAlign: 'center', fontWeight: 'bold', fontSize: '13pt' }}>
                  <th style={{ border: '1px solid black', padding: '3px 2px', width: '6%' }}>S<br />TT</th>
                  <th style={{ border: '1px solid black', padding: '3px 4px', width: '22%' }}>Chi bộ</th>
                  <th style={{ border: '1px solid black', padding: '3px 2px', width: '8%' }}>Số<br />lượng</th>
                  <th style={{ border: '1px solid black', padding: '3px 4px', width: '22%' }}>Thời<br />gian</th>
                  <th style={{ border: '1px solid black', padding: '3px 4px', width: '18%' }}>Địa điểm</th>
                  <th style={{ border: '1px solid black', padding: '3px 4px', width: '24%' }}>Bí thư,<br />điện thoại</th>
                </tr>
              </thead>
              <tbody>
                {currentMonthSchedules.map((item, idx) => {
                  const isThuyOrThuy = item.biThu && (item.biThu.includes('Trần Thị Thu Thanh Thủy') || item.biThu.includes('Nguyễn Phúc Xuân Thụy'));
                  const isThuyLong = item.biThu && item.biThu.includes('Trần Thị Thu Thanh Thủy');
                  
                  let datePart = item.thoiGian || '';
                  let dowPart = '';
                  if (item.thoiGian && item.thoiGian.includes('(')) {
                    const parts = item.thoiGian.split('(');
                    datePart = parts[0].trim();
                    dowPart = '(' + parts[1].trim();
                  }

                  return (
                    <tr key={idx} style={{ height: '22px' }}>
                      <td style={{ border: '1px solid black', padding: '2px 2px', textAlign: 'center' }}>{idx + 1}</td>
                      <td style={{ border: '1px solid black', padding: '2px 4px' }}>{item.chiBo}</td>
                      <td style={{ border: '1px solid black', padding: '2px 2px', textAlign: 'center' }}>{item.sl}</td>
                      <td style={{ border: '1px solid black', padding: '2px 3px', textAlign: 'center', lineHeight: '1.15' }}>
                        <div>{datePart}</div>
                        {dowPart && <div style={{ fontSize: '10pt' }}>{dowPart}</div>}
                      </td>
                      <td style={{ border: '1px solid black', padding: '2px 4px', textAlign: 'center' }}>{item.diaDiem}</td>
                      <td style={{ border: '1px solid black', padding: '1px 2px', textAlign: 'center', lineHeight: '1.1' }}>
                        <div style={{ 
                          fontSize: isThuyOrThuy ? '12pt' : '13pt', 
                          letterSpacing: isThuyLong ? '-0.3px' : 'normal',
                          whiteSpace: 'nowrap',
                          fontWeight: 'bold'
                        }}>
                          {item.biThu}
                        </div>
                        <div style={{ fontSize: '10.5pt' }}>{item.sdt}</div>
                      </td>
                    </tr>
                  );
                })}
                <tr style={{ fontWeight: 'bold', textAlign: 'center' }}>
                  <td colSpan="2" style={{ border: '1px solid black', padding: '3px 4px', textAlign: 'center' }}>TỔNG SỐ</td>
                  <td style={{ border: '1px solid black', padding: '3px 2px' }}>${totalMembersCount}</td>
                  <td style={{ border: '1px solid black', padding: '3px' }}></td>
                  <td style={{ border: '1px solid black', padding: '3px' }}></td>
                  <td style={{ border: '1px solid black', padding: '3px' }}></td>
                </tr>
              </tbody>
            </table>

            <table style={{ width: '100%', border: 'none', borderCollapse: 'collapse', marginTop: '4pt' }}>
              <tbody>
                <tr style={{ border: 'none' }}>
                  <td style={{ width: '50%', verticalAlign: 'top', border: 'none', padding: 0 }}>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}><u>Nơi nhận:</u></div>
                    <div style={{ fontSize: '12pt', lineHeight: '1.2' }}>
                      - Đảng ủy Sở Công Thương;<br />
                      - Bí thư các Chi bộ trực thuộc;<br />
                      - Lưu: Đảng ủy.
                    </div>
                  </td>
                  <td style={{ width: '50%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0 }}>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}>T/M ĐẢNG ỦY</div>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}>BÍ THƯ</div>
                    <div style={{ height: '38pt' }}></div>
                    <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>Nguyễn Trung Tiến</div>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* PRINT DOCUMENT 2: DANH SÁCH ĐGXL THEO HD 01-HD/TU (1 TRANG A4 Y HỆT MẪU ẢNH) */}
        {printDocType === 'dgxl' && (
          <div className="hidden print-page text-black bg-white" style={{ fontFamily: '"Times New Roman", Times, serif', width: '100%', fontSize: '13pt', lineHeight: '1.2' }}>
            <table style={{ width: '100%', border: 'none', borderCollapse: 'collapse', marginBottom: '4pt' }}>
              <tbody>
                <tr style={{ border: 'none' }}>
                  <td style={{ width: '54%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0, whiteSpace: 'nowrap' }}>
                    <div style={{ fontSize: '11pt' }}>ĐẢNG BỘ SỞ CÔNG THƯƠNG TỈNH AN GIANG</div>
                    <div style={{ fontSize: '11pt', fontWeight: 'bold' }}>ĐẢNG ỦY CHI CỤC QUẢN LÝ THỊ TRƯỜNG</div>
                    <div style={{ fontSize: '10pt', margin: '1px 0' }}>*</div>
                  </td>
                  <td style={{ width: '46%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0, whiteSpace: 'nowrap' }}>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}>ĐẢNG CỘNG SẢN VIỆT NAM</div>
                    <div style={{ height: '10pt' }}></div>
                    <div style={{ fontSize: '12pt', fontStyle: 'italic' }}>
                      An Giang, ngày 10 tháng ${String(selectedMonth).padStart(2, '0')} năm ${selectedYear}
                    </div>
                  </td>
                </tr>
              </tbody>
            </table>

            <div style={{ textAlign: 'center', margin: '4pt 0 4pt 0' }}>
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>DANH SÁCH</div>
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>
                TỔNG HỢP ĐÁNH GIÁ, XẾP LOẠI CHẤT LƯỢNG SINH HOẠT CHI BỘ
              </div>
              <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>
                Tháng ${selectedMonth} NĂM ${selectedYear}
              </div>
              <div style={{ fontSize: '11pt' }}>-----</div>
            </div>

            <table className="schedule-table" style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid black', fontSize: '10.5pt', margin: '3pt 0' }}>
              <thead>
                <tr style={{ textAlign: 'center', fontWeight: 'bold', fontSize: '11pt', backgroundColor: '#f2f2f2' }}>
                  <th rowSpan="2" style={{ border: '1px solid black', padding: '2px', width: '4%' }}>S<br/>T<br/>T</th>
                  <th rowSpan="2" style={{ border: '1px solid black', padding: '2px 4px', width: '16%' }}>Tên đơn vị</th>
                  <th rowSpan="2" style={{ border: '1px solid black', padding: '2px', width: '9%' }}>Số<br/>lượng<br/>đảng viên</th>
                  <th rowSpan="2" style={{ border: '1px solid black', padding: '2px', width: '10%' }}>Ngày<br/>họp<br/>chi<br/>bộ</th>
                  <th colSpan="5" style={{ border: '1px solid black', padding: '2px' }}>Mức đánh giá xếp loại sau buổi<br/>sinh hoạt chi bộ hằng tháng</th>
                  <th rowSpan="2" style={{ border: '1px solid black', padding: '2px 4px', width: '23%' }}>Ghi chú<br/>(thuyết minh điểm trừ)</th>
                </tr>
                <tr style={{ textAlign: 'center', fontWeight: 'bold', fontSize: '10.5pt', backgroundColor: '#f2f2f2' }}>
                  <th style={{ border: '1px solid black', padding: '2px', width: '9%' }}>KQ số<br/>điểm<br/>được<br/>đánh giá</th>
                  <th style={{ border: '1px solid black', padding: '2px', width: '7%' }}>Tốt</th>
                  <th style={{ border: '1px solid black', padding: '2px', width: '7%' }}>Khá</th>
                  <th style={{ border: '1px solid black', padding: '2px', width: '7%' }}>Trung<br/>bình</th>
                  <th style={{ border: '1px solid black', padding: '2px', width: '7%' }}>Kém</th>
                </tr>
                <tr style={{ textAlign: 'center', fontStyle: 'italic', fontSize: '9.5pt', backgroundColor: '#fafafa' }}>
                  <td style={{ border: '1px solid black', padding: '1px' }}>1</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>2</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>3</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>4</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>5</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>6</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>7</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>8</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>9</td>
                  <td style={{ border: '1px solid black', padding: '1px' }}>10</td>
                </tr>
              </thead>
              <tbody>
                {currentMonthDGXL.map((item, idx) => (
                  <tr key={idx} style={{ height: '20px' }}>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{idx + 1}</td>
                    <td style={{ border: '1px solid black', padding: '2px 4px' }}>{item.chiBo}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.sl || ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.ngayHop || ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.diemDG !== '' ? item.diemDG : ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.mucXepLoai === 'Tốt' ? 'x' : ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.mucXepLoai === 'Khá' ? 'x' : ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.mucXepLoai === 'Trung bình' ? 'x' : ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px', textAlign: 'center' }}>{item.mucXepLoai === 'Kém' ? 'x' : ''}</td>
                    <td style={{ border: '1px solid black', padding: '2px 4px', fontSize: '10pt' }}>{item.ghiChuTruDiem || ''}</td>
                  </tr>
                ))}
                <tr style={{ fontWeight: 'bold', textAlign: 'center', height: '20px' }}>
                  <td colSpan="2" style={{ border: '1px solid black', padding: '2px 4px', textAlign: 'center' }}>Tổng cộng</td>
                  <td style={{ border: '1px solid black', padding: '2px' }}></td>
                  <td style={{ border: '1px solid black', padding: '2px' }}></td>
                  <td style={{ border: '1px solid black', padding: '2px' }}></td>
                  <td style={{ border: '1px solid black', padding: '2px' }}>{currentMonthDGXL.filter(d => d.mucXepLoai === 'Tốt').length || ''}</td>
                  <td style={{ border: '1px solid black', padding: '2px' }}>{currentMonthDGXL.filter(d => d.mucXepLoai === 'Khá').length || ''}</td>
                  <td style={{ border: '1px solid black', padding: '2px' }}>{currentMonthDGXL.filter(d => d.mucXepLoai === 'Trung bình').length || ''}</td>
                  <td style={{ border: '1px solid black', padding: '2px' }}>{currentMonthDGXL.filter(d => d.mucXepLoai === 'Kém').length || ''}</td>
                  <td style={{ border: '1px solid black', padding: '2px' }}></td>
                </tr>
              </tbody>
            </table>

            {/* Dòng thuyết minh biến động */}
            <div style={{ fontSize: '11pt', marginTop: '3pt', marginBottom: '6pt', fontStyle: 'italic' }}>
              {currentMonthNote}
            </div>

            <table style={{ width: '100%', border: 'none', borderCollapse: 'collapse', marginTop: '4pt' }}>
              <tbody>
                <tr style={{ border: 'none' }}>
                  <td style={{ width: '50%', verticalAlign: 'top', border: 'none', padding: 0 }}>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}><u>Nơi nhận:</u></div>
                    <div style={{ fontSize: '12pt', lineHeight: '1.2' }}>
                      - Đảng ủy Sở;<br />
                      - Đảng ủy Chi cục QLTT,<br />
                      - Lưu: ĐU.
                    </div>
                  </td>
                  <td style={{ width: '50%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0 }}>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}>T/M ĐẢNG ỦY</div>
                    <div style={{ fontSize: '12pt', fontWeight: 'bold' }}>BÍ THƯ</div>
                    <div style={{ height: '38pt' }}></div>
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
                    Đăng ký trước ngày 20 • Thống nhất lãnh đạo ngày 25 • Ban hành thông báo chính thức
                  </p>
                </div>

                <div className="flex items-center gap-3">
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
                        <th className="w-48">Chi bộ</th>
                        <th className="w-16 text-center">Số lượng</th>
                        <th className="w-48">Thời gian sinh hoạt</th>
                        <th className="w-44">Địa điểm</th>
                        <th className="w-48">Bí thư, điện thoại</th>
                        <th>Ghi chú biến động</th>
                        <th className="w-32 text-center">Trạng thái</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentMonthSchedules.map((item, idx) => (
                        <tr key={item.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{item.chiBo}</td>
                          <td className="text-center font-bold text-red-600 text-xs">{item.sl}</td>
                          <td>
                            <input
                              type="text"
                              value={item.thoiGian}
                              onChange={(e) => handleUpdateSchedule(item.chiBoId, 'thoiGian', e.target.value)}
                              placeholder="VD: 03/9/2026 (thứ Năm)"
                              className="text-xs font-medium text-blue-600 dark:text-blue-400"
                            />
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
                    Bảng 10 cột chuẩn theo Hướng dẫn 01-HD/TU. Sau khi các Đội nhập xong qua link, bấm <b>"Xuất File Word (.DOC)"</b> hoặc <b>"In Báo Cáo (1 Trang)"</b> để trình ký Bí thư.
                  </p>
                </div>

                <div className="flex items-center gap-3">
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
                    className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer transition-all"
                  >
                    <FileSpreadsheet className="w-4 h-4" /> Xuất Excel
                  </button>
                </div>
              </div>

              {/* BẢNG 10 CỘT THEO HƯỚNG DẪN 01-HD/TU (Y HỆT HÌNH ẢNH MẪU) */}
              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th rowSpan="2" className="w-10 text-center">S<br/>T<br/>T</th>
                        <th rowSpan="2" className="w-40">Tên đơn vị</th>
                        <th rowSpan="2" className="w-28 text-center">Số<br/>lượng<br/>đảng viên</th>
                        <th rowSpan="2" className="w-28 text-center">Ngày<br/>họp<br/>chi<br/>bộ</th>
                        <th colSpan="5" className="text-center bg-blue-50/50 dark:bg-blue-950/30">
                          Mức đánh giá xếp loại sau buổi sinh hoạt chi bộ hằng tháng
                        </th>
                        <th rowSpan="2" className="w-64">Ghi chú<br/>(thuyết minh điểm trừ)</th>
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
                            <input
                              type="text"
                              value={d.ngayHop || ''}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'ngayHop', e.target.value)}
                              placeholder="03/9/2026"
                              className="w-24 text-center text-xs font-medium text-blue-600 dark:text-blue-400"
                            />
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
                              placeholder="Thuyết minh lý do trừ điểm hoặc ghi chú..."
                              value={d.ghiChuTruDiem || ''}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'ghiChuTruDiem', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                        </tr>
                      ))}
                      <tr className="font-bold bg-slate-100/70 dark:bg-slate-800/70">
                        <td colSpan="2" className="text-center text-xs uppercase font-bold">Tổng cộng</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td className="text-center text-xs text-emerald-600 font-bold">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Tốt').length || ''}</td>
                        <td className="text-center text-xs text-blue-600 font-bold">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Khá').length || ''}</td>
                        <td className="text-center text-xs text-amber-600 font-bold">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Trung bình').length || ''}</td>
                        <td className="text-center text-xs text-red-600 font-bold">{currentMonthDGXL.filter(d => d.mucXepLoai === 'Kém').length || ''}</td>
                        <td></td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Dòng ghi chú dưới chân bảng */}
                <div className="p-4 border-t border-[var(--border-color)] bg-slate-50/50 dark:bg-slate-900/30">
                  <label className="block text-xs font-bold text-slate-700 dark:text-slate-300 mb-1">
                    Ghi chú biến động đảng viên chân trang (In kèm trong báo cáo):
                  </label>
                  <input
                    type="text"
                    value={currentMonthNote}
                    onChange={(e) => handleUpdateOverallNote(e.target.value)}
                    className="w-full text-xs font-medium italic text-slate-800 dark:text-slate-200"
                  />
                </div>
              </div>
            </div>
          )}

          {/* TAB 4: BIỂU TỔNG HỢP SỐ LIỆU BÁO CÁO 06 THÁNG NĂM 2026 */}
          {activeTab === 'sixmonths' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <h3 className="text-base font-bold text-[var(--text-main)]">
                    Biểu Tổng Hợp Số Liệu Báo Cáo 06 Tháng Năm {selectedYear}
                  </h3>
                  <p className="text-xs text-[var(--text-muted)]">Tổng hợp ĐV dự sinh hoạt, phát biểu, Dân vận khéo, Chuyên đề, Kiểm tra giám sát của 13 Chi bộ</p>
                </div>

                <div className="flex items-center gap-3">
                  <button 
                    onClick={handleExportSixMonthsToExcel}
                    className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5" /> Xuất File Báo Cáo 06 Tháng (Excel)
                  </button>
                </div>
              </div>

              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th className="w-10 text-center">STT</th>
                        <th className="w-32">Tên chi bộ</th>
                        <th className="w-16 text-center">Đảng số</th>
                        <th className="w-36">Số lượng ĐV dự sinh hoạt lệ</th>
                        <th className="w-36">Phát biểu ý kiến</th>
                        <th className="w-48">Dân vận khéo</th>
                        <th className="w-48">Học tập Chuyên đề {selectedYear}</th>
                        <th className="w-48">Đọc bài viết</th>
                        <th className="w-48">KH Sinh hoạt CĐ</th>
                        <th className="w-16 text-center">KNĐ</th>
                        <th className="w-48">KT GS</th>
                        <th>Ghi chú</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sixMonthsData.map((d, idx) => (
                        <tr key={d.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{d.chiBo}</td>
                          <td className="text-center font-bold text-red-600 text-xs">{d.dangSo}</td>
                          <td>
                            <input
                              type="text"
                              value={d.soLgDuHop}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'soLgDuHop', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.phatBieuYkien}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'phatBieuYkien', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.danVanKheo}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'danVanKheo', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.hocTapChuyenDe}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'hocTapChuyenDe', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.docBaiViet}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'docBaiViet', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.khSinhHoatCD}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'khSinhHoatCD', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td className="text-center">
                            <input
                              type="text"
                              value={d.knd}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'knd', e.target.value)}
                              className="w-12 text-center text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.ktgs}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'ktgs', e.target.value)}
                              className="text-xs"
                            />
                          </td>
                          <td>
                            <input
                              type="text"
                              value={d.ghiChu}
                              onChange={(e) => handleUpdateSixMonths(d.chiBoId, 'ghiChu', e.target.value)}
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

          {/* TAB 5: DASHBOARD */}
          {activeTab === 'dashboard' && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="card-glass p-5 flex items-center justify-between border-l-4 border-red-500">
                  <div>
                    <p className="text-xs font-medium text-[var(--text-muted)] uppercase">Tổng số Đảng viên</p>
                    <h3 className="text-3xl font-extrabold mt-1 text-[var(--text-main)]">{totalMembersCount}</h3>
                    <p className="text-xs text-emerald-600 mt-1 flex items-center gap-1 font-medium">
                      <CheckCircle2 className="w-3.5 h-3.5" /> 13/13 Chi bộ trực thuộc
                    </p>
                  </div>
                  <div className="w-12 h-12 rounded-xl bg-red-50 dark:bg-red-950/50 flex items-center justify-center text-red-600">
                    <Users className="w-6 h-6" />
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center justify-between border-l-4 border-blue-500">
                  <div>
                    <p className="text-xs font-medium text-[var(--text-muted)] uppercase">Chính thức</p>
                    <h3 className="text-3xl font-extrabold mt-1 text-[var(--text-main)]">152</h3>
                    <p className="text-xs text-[var(--text-muted)] mt-1">Tỷ lệ: 97%</p>
                  </div>
                  <div className="w-12 h-12 rounded-xl bg-blue-50 dark:bg-blue-950/50 flex items-center justify-center text-blue-600">
                    <ShieldCheck className="w-6 h-6" />
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center justify-between border-l-4 border-amber-500">
                  <div>
                    <p className="text-xs font-medium text-[var(--text-muted)] uppercase">Dự bị</p>
                    <h3 className="text-3xl font-extrabold mt-1 text-amber-600">5</h3>
                    <p className="text-xs text-amber-600 mt-1 flex items-center gap-1">
                      <Clock className="w-3.5 h-3.5" /> Theo dõi rèn luyện
                    </p>
                  </div>
                  <div className="w-12 h-12 rounded-xl bg-amber-50 dark:bg-amber-950/50 flex items-center justify-center text-amber-600">
                    <Clock className="w-6 h-6" />
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center justify-between border-l-4 border-purple-500">
                  <div>
                    <p className="text-xs font-medium text-[var(--text-muted)] uppercase">Đảng viên Nữ</p>
                    <h3 className="text-3xl font-extrabold mt-1 text-[var(--text-main)]">42</h3>
                    <p className="text-xs text-[var(--text-muted)] mt-1">Tỷ lệ: 27%</p>
                  </div>
                  <div className="w-12 h-12 rounded-xl bg-purple-50 dark:bg-purple-950/50 flex items-center justify-center text-purple-600">
                    <Award className="w-6 h-6" />
                  </div>
                </div>
              </div>

              {/* DANH SÁCH 13 CHI BỘ */}
              <div className="card-glass p-5">
                <h3 className="text-base font-bold text-[var(--text-main)] mb-1">Cơ cấu 13 Chi bộ trực thuộc & Bí thư</h3>
                <div className="table-container mt-4">
                  <table>
                    <thead>
                      <tr>
                        <th>STT</th>
                        <th>Tên Chi bộ</th>
                        <th className="text-center">Số lượng ĐV</th>
                        <th>Địa bàn / Địa điểm</th>
                        <th>Bí thư Chi bộ</th>
                        <th>Điện thoại</th>
                      </tr>
                    </thead>
                    <tbody>
                      {branchDetails.map((b, idx) => (
                        <tr key={b.id}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{b.name}</td>
                          <td className="text-center font-bold text-red-600">{b.sl}</td>
                          <td className="text-xs">{b.diaDiem}</td>
                          <td className="font-semibold text-xs">{b.biThu}</td>
                          <td className="font-mono text-xs text-[var(--text-muted)]">{b.sdt}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* TAB 6: MEMBERS */}
          {activeTab === 'members' && (
            <div className="space-y-4">
              <div className="card-glass p-4 flex flex-col md:flex-row items-center justify-between gap-4">
                <div className="flex flex-1 items-center gap-3 w-full">
                  <div className="relative flex-1">
                    <Search className="w-4 h-4 absolute left-3 top-2.5 text-[var(--text-muted)]" />
                    <input
                      type="text"
                      placeholder="Tìm kiếm theo tên, số thẻ đảng, số lý lịch..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="pl-9 text-xs"
                    />
                  </div>

                  <div className="w-48">
                    <select
                      value={filterChiBo}
                      onChange={(e) => setFilterChiBo(e.target.value)}
                      className="text-xs"
                    >
                      <option value="Tất cả">Tất cả Chi bộ</option>
                      {chiBoList.map(cb => <option key={cb} value={cb}>{cb}</option>)}
                    </select>
                  </div>
                </div>
              </div>

              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th>STT</th>
                        <th>Họ và Tên</th>
                        <th>Chi bộ</th>
                        <th>Chức vụ Đảng/Chính quyền</th>
                        <th>Số Thẻ / Lý lịch</th>
                        <th>Ngày vào Đảng</th>
                        <th>Loại ĐV</th>
                        <th>Lý luận chính trị</th>
                      </tr>
                    </thead>
                    <tbody>
                      {members.map((m, idx) => (
                        <tr key={m.id}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td>
                            <div className="font-bold text-[var(--text-main)]">{m.hoTen}</div>
                            <div className="text-xs text-[var(--text-muted)]">{m.gioiTinh} - {m.ngaySinh}</div>
                          </td>
                          <td className="text-xs font-medium">{m.chiBo}</td>
                          <td className="text-xs">
                            <div className="font-medium text-red-600 dark:text-red-400">{m.chucVuDang}</div>
                            <div className="text-[var(--text-muted)]">{m.chucVuChinhQuyen}</div>
                          </td>
                          <td className="text-xs font-mono">
                            <div>Thẻ: {m.soTheDang || '-'}</div>
                            <div className="text-[var(--text-muted)]">LL: {m.soLyLich || '-'}</div>
                          </td>
                          <td className="text-xs">{m.ngayVaoDang}</td>
                          <td>
                            <span className="badge bg-blue-50 text-blue-700 border border-blue-200">
                              {m.loaiDangVien}
                            </span>
                          </td>
                          <td className="text-xs font-medium">{m.lyLuanChinhTri}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

        </main>

        {/* MODAL: CHIA SẺ LINK VÀ MẪU TIN NHẮN ZALO CHO CHI BỘ */}
        {isShareModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-[var(--bg-card)] border border-[var(--border-color)] rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl animate-in fade-in zoom-in-95 duration-150">
              <div className="p-5 border-b border-[var(--border-color)] flex items-center justify-between bg-indigo-50/50 dark:bg-indigo-950/20">
                <div className="flex items-center gap-2">
                  <Share2 className="w-5 h-5 text-indigo-600" />
                  <h3 className="font-bold text-base text-[var(--text-main)]">
                    Gửi Link & Mẫu Tin Nhắn Zalo Cho Các Chi Bộ
                  </h3>
                </div>
                <button onClick={() => setIsShareModalOpen(false)} className="text-lg font-bold cursor-pointer">✕</button>
              </div>

              {/* Share Tabs */}
              <div className="flex border-b border-[var(--border-color)] px-6 pt-3 bg-slate-50 dark:bg-slate-900/50 gap-4">
                <button
                  onClick={() => setShareTab('register')}
                  className={`pb-2.5 text-xs font-bold border-b-2 cursor-pointer transition-all ${
                    shareTab === 'register' 
                      ? 'border-red-600 text-red-600 dark:text-red-400' 
                      : 'border-transparent text-[var(--text-muted)] hover:text-[var(--text-main)]'
                  }`}
                >
                  📅 1. Đăng ký Lịch họp (Trước ngày 20)
                </button>
                <button
                  onClick={() => setShareTab('report')}
                  className={`pb-2.5 text-xs font-bold border-b-2 cursor-pointer transition-all ${
                    shareTab === 'report' 
                      ? 'border-blue-600 text-blue-600 dark:text-blue-400' 
                      : 'border-transparent text-[var(--text-muted)] hover:text-[var(--text-main)]'
                  }`}
                >
                  ⭐ 2. Báo cáo ĐGXL sau sinh hoạt (HD 01-HD/TU)
                </button>
              </div>

              <div className="p-6 space-y-4">
                {shareTab === 'register' ? (
                  <>
                    <div>
                      <label className="block text-xs font-bold mb-1">Link trực tiếp gửi cho các Đội tự chọn ngày họp:</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          readOnly
                          value={window.location.origin + window.location.pathname + '?mode=register'}
                          className="text-xs font-mono bg-slate-50 dark:bg-slate-900"
                        />
                        <button
                          onClick={() => {
                            navigator.clipboard.writeText(window.location.origin + window.location.pathname + '?mode=register');
                            setCopiedLink(true);
                            setTimeout(() => setCopiedLink(false), 2000);
                          }}
                          className="px-3 py-2 bg-indigo-600 text-white rounded-md text-xs font-bold flex items-center gap-1 shrink-0 cursor-pointer"
                        >
                          {copiedLink ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                          {copiedLink ? 'Đã chép' : 'Chép Link'}
                        </button>
                      </div>
                    </div>

                    <div>
                      <label className="block text-xs font-bold mb-1">Mẫu tin nhắn soạn sẵn gửi nhóm Zalo các Bí thư / Đội trưởng:</label>
                      <textarea
                        rows={5}
                        readOnly
                        value={sampleZaloRegisterMessage}
                        className="text-xs font-sans bg-slate-50 dark:bg-slate-900 p-3 leading-relaxed"
                      />
                      <button
                        onClick={() => {
                          navigator.clipboard.writeText(sampleZaloRegisterMessage);
                          setCopiedMessage(true);
                          setTimeout(() => setCopiedMessage(false), 2000);
                        }}
                        className="mt-2 w-full py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-md flex items-center justify-center gap-1 cursor-pointer"
                      >
                        {copiedMessage ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                        {copiedMessage ? 'Đã sao chép tin nhắn Zalo' : 'Sao chép Toàn bộ Tin Nhắn gửi Zalo'}
                      </button>
                    </div>
                  </>
                ) : (
                  <>
                    <div>
                      <label className="block text-xs font-bold mb-1">Link trực tiếp gửi cho các Chi bộ gửi Báo cáo ĐGXL sau họp:</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          readOnly
                          value={window.location.origin + window.location.pathname + '?mode=report'}
                          className="text-xs font-mono bg-slate-50 dark:bg-slate-900"
                        />
                        <button
                          onClick={() => {
                            navigator.clipboard.writeText(window.location.origin + window.location.pathname + '?mode=report');
                            setCopiedLink(true);
                            setTimeout(() => setCopiedLink(false), 2000);
                          }}
                          className="px-3 py-2 bg-blue-600 text-white rounded-md text-xs font-bold flex items-center gap-1 shrink-0 cursor-pointer"
                        >
                          {copiedLink ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                          {copiedLink ? 'Đã chép' : 'Chép Link'}
                        </button>
                      </div>
                    </div>

                    <div>
                      <label className="block text-xs font-bold mb-1">Mẫu tin nhắn nhắc Báo cáo ĐGXL gửi nhóm Zalo:</label>
                      <textarea
                        rows={5}
                        readOnly
                        value={sampleZaloReportMessage}
                        className="text-xs font-sans bg-slate-50 dark:bg-slate-900 p-3 leading-relaxed"
                      />
                      <button
                        onClick={() => {
                          navigator.clipboard.writeText(sampleZaloReportMessage);
                          setCopiedMessage(true);
                          setTimeout(() => setCopiedMessage(false), 2000);
                        }}
                        className="mt-2 w-full py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-md flex items-center justify-center gap-1 cursor-pointer"
                      >
                        {copiedMessage ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                        {copiedMessage ? 'Đã sao chép tin nhắn Zalo' : 'Sao chép Toàn bộ Tin Nhắn gửi Zalo'}
                      </button>
                    </div>
                  </>
                )}

                <div className="p-3 rounded-lg bg-blue-50 dark:bg-blue-950/40 text-xs text-blue-800 dark:text-blue-200">
                  💡 <b>Tiện ích:</b> Chi bộ mở link trên điện thoại hoặc máy tính, nhập điểm tự chấm và số lượng ĐV là số liệu sẽ lập tức được tổng hợp tự động vào Biểu mẫu ĐGXL của Đảng ủy!
                </div>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
