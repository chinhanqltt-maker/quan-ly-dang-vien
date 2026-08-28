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
  TrendingUp,
  Calendar,
  Award,
  CreditCard,
  FileSpreadsheet,
  Printer,
  Plus,
  Search,
  Filter,
  Trash2,
  Edit,
  CheckCircle2,
  AlertCircle,
  Clock,
  ShieldCheck,
  ChevronRight,
  Sun,
  Moon,
  Sparkles,
  Download,
  Upload,
  UserPlus,
  Send,
  CheckCheck,
  Share2,
  Copy,
  Check,
  MapPin,
  Archive,
  Save,
  BookOpen,
  FileCheck2,
  FileText
} from 'lucide-react';

export default function App() {
  const [darkMode, setDarkMode] = useState(false);

  // Check URL params for Member Registration Mode
  const urlParams = new URLSearchParams(window.location.search);
  const isPublicRegisterMode = urlParams.get('mode') === 'register';
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

  // Biểu Đánh Giá Xếp Loại Chất Lượng Sinh Hoạt Chi Bộ Hằng Tháng (ĐGXL)
  const [dgxlData, setDgxlData] = useState(() => {
    const saved = localStorage.getItem('qltt_party_dgxl_v3');
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
  const [selectedMonth, setSelectedMonth] = useState(8);

  const [searchTerm, setSearchTerm] = useState('');
  const [filterChiBo, setFilterChiBo] = useState('Tất cả');

  // Modal Share Link
  const [isShareModalOpen, setIsShareModalOpen] = useState(false);
  const [copiedLink, setCopiedLink] = useState(false);
  const [copiedMessage, setCopiedMessage] = useState(false);

  // Form tự đăng ký của Đội
  const initialSelectedBranchId = paramChiBoId ? Number(paramChiBoId) : 1;
  const initialBranchObj = branchDetails.find(b => b.id === initialSelectedBranchId) || branchDetails[0];

  const [teamSelectChiBo, setTeamSelectChiBo] = useState(initialSelectedBranchId);
  const [teamDate, setTeamDate] = useState('');
  const [teamDayOfWeek, setTeamDayOfWeek] = useState('thứ Hai');
  const [teamLocation, setTeamLocation] = useState(initialBranchObj.diaDiem);
  const [teamNote, setTeamNote] = useState('');
  const [teamSubmitted, setTeamSubmitted] = useState(false);

  useEffect(() => {
    const br = branchDetails.find(b => b.id === teamSelectChiBo);
    if (br) {
      setTeamLocation(br.diaDiem);
    }
  }, [teamSelectChiBo, branchDetails]);

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
    localStorage.setItem('qltt_party_sixmonths_v3', JSON.stringify(sixMonthsReports));
  }, [sixMonthsReports]);
  useEffect(() => {
    localStorage.setItem('qltt_party_archived_notices', JSON.stringify(archivedNotices));
  }, [archivedNotices]);

  const currentKey = `${selectedYear}-${selectedMonth}`;

  const currentMonthSchedules = useMemo(() => {
    if (meetingSchedules[currentKey]) {
      // Đảm bảo luôn giữ đúng Bí thư, Số điện thoại và Địa điểm từ danh mục chuẩn nếu bị thiếu
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

    alert(`Đã duyệt thống nhất với Lãnh đạo và tự động lưu vào 'Kho Lưu Trữ Thông Báo'! Đồng chí có thể bấm 'Xuất File Word (.DOCX)' hoặc 'In Trình Ký'.`);
  };

  const currentMonthDGXL = useMemo(() => {
    if (dgxlData[currentKey]) {
      return dgxlData[currentKey];
    }
    return branchDetails.map(b => {
      const schedule = currentMonthSchedules.find(s => s.chiBoId === b.id) || {};
      return {
        chiBoId: b.id,
        chiBo: b.name,
        sl: b.sl,
        ngayHop: schedule.thoiGian || 'Chưa họp',
        diemDG: 100,
        mucXepLoai: 'Tốt',
        ghiChuTruDiem: schedule.ghiChu || ''
      };
    });
  }, [dgxlData, currentKey, branchDetails, currentMonthSchedules]);

  const handleUpdateDGXL = (chiBoId, field, value) => {
    const updated = currentMonthDGXL.map(item => {
      if (item.chiBoId === chiBoId) {
        return { ...item, [field]: value };
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

  const handleExportDGXLToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(currentMonthDGXL.map((d, idx) => ({
      'STT': idx + 1,
      'Tên đơn vị': d.chiBo,
      'Số lượng đảng viên': d.sl,
      'Ngày họp chi bộ': d.ngayHop,
      'Kết quả số điểm đánh giá': d.diemDG,
      'Tốt': d.mucXepLoai === 'Tốt' ? 'X' : '',
      'Khá': d.mucXepLoai === 'Khá' ? 'X' : '',
      'Trung bình': d.mucXepLoai === 'Trung bình' ? 'X' : '',
      'Kém': d.mucXepLoai === 'Kém' ? 'X' : '',
      'Ghi chú (thuyết minh điểm trừ)': d.ghiChuTruDiem
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `DGXL_T${selectedMonth}_${selectedYear}`);
    XLSX.writeFile(wb, `DGXL_Sinh_Hoat_Chi_Bo_T${selectedMonth}_${selectedYear}.xlsx`);
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

  // XUẤT FILE WORD .DOC CHUẨN ĐỊNH DẠNG MICROSOFT WORD 100% KHÔNG LỖI
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

  const handlePrint = () => {
    const originalTitle = document.title;
    document.title = "";
    window.print();
    setTimeout(() => {
      document.title = originalTitle;
    }, 1000);
  };

  const sampleZaloMessage = `[THÔNG BÁO ĐẢNG ỦY BỘ PHẬN CHI CỤC QLTT]
Kính gửi: Bí thư các Chi bộ trực thuộc (Đội 1 đến Đội 12 và Khối phòng).
Thực hiện Quy chế làm việc, đề nghị các Chi bộ chủ động đăng ký lịch sinh hoạt lệ tháng ${selectedMonth}/${selectedYear} trước ngày 20 để Đảng ủy tổng hợp xin ý kiến Lãnh đạo và ban hành Thông báo chính thức.
👉 Link đăng ký trực tuyến: ${window.location.origin + window.location.pathname}?mode=register`;

  // GIAO DIỆN DÀNH CHO CÁC ĐỘI TỰ ĐĂNG KÝ
  if (isPublicRegisterMode) {
    const selectedBranchInfo = branchDetails.find(b => b.id === teamSelectChiBo) || branchDetails[0];

    return (
      <div className="min-h-screen bg-slate-100 flex items-center justify-center p-4">
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
                          const [yyyy, mm, dd] = e.target.value.split('-');
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

            <div className="text-center pt-2">
              <a 
                href={window.location.pathname}
                className="text-xs text-slate-500 hover:text-red-600 font-medium"
              >
                ← Quay lại Bảng Quản Trị Tổng Hợp
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
                    ĐẢNG ỦY BỘ PHẬN CHI CỤC QUẢN LÝ THỊ TRƯỜNG AN GIANG
                  </h1>
                  <span className="badge bg-red-50 text-red-700 border border-red-200 dark:bg-red-950/40 dark:text-red-300 dark:border-red-800">
                    Cán bộ tổng hợp
                  </span>
                </div>
                <p className="text-xs text-[var(--text-muted)]">Quy trình Đăng ký ngày 20, Duyệt ngày 25, Lưu thông báo, ĐGXL hằng tháng & Báo cáo 06 Tháng</p>
              </div>
            </div>

            <div className="flex items-center gap-2.5">
              <button 
                onClick={handleExportWordDoc}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-blue-700 hover:bg-blue-800 text-white shadow-sm transition-all cursor-pointer"
                title="Tải về file Word (.DOC) chuẩn 100% mở bằng Microsoft Word không bao giờ lỗi"
              >
                <FileText className="w-4 h-4" />
                Tải File Word (.DOC)
              </button>

              <button 
                onClick={() => setIsShareModalOpen(true)}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white shadow-sm transition-all cursor-pointer"
                title="Tạo thông báo và lấy link gửi cho các Đội tự đăng ký ngày họp"
              >
                <Share2 className="w-4 h-4" />
                Gửi Link Đăng Ký (Ngày 20)
              </button>

              <button 
                onClick={handlePrint}
                className="inline-flex items-center gap-1.5 px-3.5 py-2 text-xs font-bold rounded-lg bg-red-600 hover:bg-red-700 text-white shadow-sm transition-all cursor-pointer"
                title="In thông báo lịch sinh hoạt trình ký Bí thư Đảng ủy (Chuẩn khít 1 trang A4)"
              >
                <Printer className="w-4 h-4" />
                In Thông Báo (1 Trang)
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
              { id: 'dgxl', label: '⭐ 3. Đánh Giá Xếp Loại Hằng Tháng (ĐGXL)', icon: FileCheck2 },
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

        {/* PRINT DOCUMENT FORMAT - CỠ CHỮ CĂN CỨ 14PT CHUẨN THỂ THỨC VĂN BẢN ĐẢNG, KHÍT 1 TRANG A4 */}
        <div className="hidden print-page text-black bg-white" style={{ fontFamily: '"Times New Roman", Times, serif', width: '100%', fontSize: '13pt', lineHeight: '1.2' }}>
          {/* Header 2 columns */}
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

          {/* Title */}
          <div style={{ textAlign: 'center', margin: '4pt 0 2pt 0' }}>
            <div style={{ fontSize: '14pt', fontWeight: 'bold' }}>THÔNG BÁO</div>
            <div style={{ fontSize: '14pt', fontWeight: 'bold' }}>
              Lịch sinh hoạt lệ tháng {String(selectedMonth).padStart(2, '0')}/{selectedYear} của các chi bộ trực thuộc
            </div>
            <div style={{ fontSize: '12pt' }}>-----</div>
          </div>

          {/* Paragraphs Cỡ chữ 14pt chuẩn theo yêu cầu */}
          <div style={{ textIndent: '30pt', textAlign: 'justify', marginBottom: '2pt', fontSize: '14pt', lineHeight: '1.25' }}>
            Căn cứ Quy chế làm việc của Đảng ủy Chi cục Quản lý thị trường và Quy chế làm việc của các chi bộ trực thuộc nhiệm kỳ 2025-2030.
          </div>
          <div style={{ textIndent: '30pt', textAlign: 'justify', marginBottom: '4pt', fontSize: '14pt', lineHeight: '1.25' }}>
            Theo đăng ký lịch sinh hoạt lệ chi bộ tháng {String(selectedMonth).padStart(2, '0')}/{selectedYear}. Đảng ủy bộ phận Chi cục Quản lý thị trường thông báo thời gian, địa điểm sinh hoạt của các chi bộ, như sau:
          </div>

          {/* 13 Branch Schedule Table */}
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
                
                // Tách Ngày và Thứ thành 2 dòng riêng biệt
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
                <td style={{ border: '1px solid black', padding: '3px 2px' }}>{totalMembersCount}</td>
                <td style={{ border: '1px solid black', padding: '3px' }}></td>
                <td style={{ border: '1px solid black', padding: '3px' }}></td>
                <td style={{ border: '1px solid black', padding: '3px' }}></td>
              </tr>
            </tbody>
          </table>

          {/* Footer Table */}
          <table style={{ width: '100%', border: 'none', borderCollapse: 'collapse', marginTop: '4pt' }}>
            <tbody>
              <tr style={{ border: 'none' }}>
                <td style={{ width: '50%', verticalAlign: 'top', border: 'none', padding: 0 }}>
                  <div style={{ fontSize: '12pt', fontWeight: 'bold', textDecoration: 'underline' }}><u>Nơi nhận:</u></div>
                  <div style={{ fontSize: '12pt', lineHeight: '1.25' }}>
                    - Đảng ủy Sở Công Thương;<br />
                    - Bí thư các Chi bộ trực thuộc;<br />
                    - Lưu: Đảng ủy.
                  </div>
                </td>
                <td style={{ width: '50%', textAlign: 'center', verticalAlign: 'top', border: 'none', padding: 0 }}>
                  <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>T/M ĐẢNG ỦY</div>
                  <div style={{ fontSize: '13pt', fontWeight: 'bold' }}>BÍ THƯ</div>
                  <div style={{ height: '40pt' }}></div>
                  <div style={{ fontSize: '14pt', fontWeight: 'bold' }}>Nguyễn Trung Tiến</div>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* MAIN BODY CONTENT */}
        <main className="max-w-7xl mx-auto p-6 flex-1 w-full no-print">

          {/* TAB 1: LỊCH HỌP & DUYỆT TRÌNH KÝ */}
          {activeTab === 'meetings' && (
            <div className="space-y-6">
              
              {/* QUY TRÌNH BANNER */}
              <div className="card-glass p-5 border-l-4 border-amber-500 bg-linear-to-r from-amber-500/5 to-transparent">
                <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                  <div>
                    <div className="flex items-center gap-2">
                      <Clock className="w-5 h-5 text-amber-600" />
                      <h3 className="text-base font-bold text-[var(--text-main)]">
                        Quy trình Quản lý Lịch họp Chi bộ định kỳ (Ngày 20 - 25 hàng tháng)
                      </h3>
                    </div>
                    <p className="text-xs text-[var(--text-muted)] mt-1">
                      1. Ngày 20 gửi link Zalo cho các Đội $ightarrow$ 2. Đội tự chọn ngày $ightarrow$ 3. Trước ngày 25 báo cáo Lãnh đạo $ightarrow$ 4. Bấm <b>'Lãnh đạo đã Thống nhất & Lưu File'</b> $ightarrow$ 5. Bấm <b>'Tải File Word'</b> hoặc <b>'In Thông Báo (1 Trang)'</b>.
                    </p>
                  </div>

                  <div className="flex items-center gap-3">
                    <div className="flex items-center gap-2">
                      <span className="text-xs font-semibold">Tháng:</span>
                      <select 
                        value={selectedMonth} 
                        onChange={(e) => setSelectedMonth(Number(e.target.value))}
                        className="w-24 text-xs font-bold"
                      >
                        {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <option key={m} value={m}>Tháng {m}</option>)}
                      </select>
                    </div>

                    <div className="flex items-center gap-2">
                      <span className="text-xs font-semibold">Năm:</span>
                      <select 
                        value={selectedYear} 
                        onChange={(e) => setSelectedYear(Number(e.target.value))}
                        className="w-24 text-xs font-bold"
                      >
                        {[2025, 2026, 2027].map(y => <option key={y} value={y}>Năm {y}</option>)}
                      </select>
                    </div>
                  </div>
                </div>
              </div>

              {/* ACTION BAR */}
              <div className="card-glass p-4 flex flex-wrap items-center justify-between gap-4">
                <div className="flex items-center gap-3">
                  <span className="badge bg-blue-50 text-blue-700 border border-blue-200">
                    Đã đăng ký: {currentMonthSchedules.filter(i => i.thoiGian && i.thoiGian !== 'Chưa đăng ký').length}/13 Chi bộ
                  </span>
                  <span className="badge bg-emerald-50 text-emerald-700 border border-emerald-200">
                    Tổng số đảng viên: {totalMembersCount} ĐV
                  </span>
                </div>

                <div className="flex items-center gap-2.5">
                  <button
                    onClick={() => setIsShareModalOpen(true)}
                    className="px-3.5 py-2 rounded-md bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                  >
                    <Share2 className="w-4 h-4" />
                    Gửi Link Zalo (Ngày 20)
                  </button>

                  <button
                    onClick={handleApproveAndArchive}
                    className="px-3.5 py-2 rounded-md bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                    title="Lãnh đạo đồng ý thống nhất và tự động lưu vào Kho lưu trữ để sau này coi lại"
                  >
                    <Save className="w-4 h-4" />
                    Lãnh đạo Thống nhất & Lưu File
                  </button>

                  <button
                    onClick={handleExportWordDoc}
                    className="px-3.5 py-2 rounded-md bg-blue-700 hover:bg-blue-800 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                    title="Xuất file Word .DOC chuẩn 100% mở bằng MS Word không bao giờ bị lỗi"
                  >
                    <FileText className="w-4 h-4" />
                    Tải File Word (.DOC)
                  </button>

                  <button
                    onClick={handlePrint}
                    className="px-3.5 py-2 rounded-md bg-red-600 hover:bg-red-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                    title="In chuẩn khít 1 trang A4"
                  >
                    <Printer className="w-4 h-4" />
                    In Thông Báo (1 Trang)
                  </button>
                </div>
              </div>

              {/* SCHEDULE TABLE */}
              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th className="w-12 text-center">STT</th>
                        <th>Chi bộ</th>
                        <th className="w-16 text-center">Số lượng ĐV</th>
                        <th className="w-60">Thời gian sinh hoạt (Đăng ký)</th>
                        <th>Địa điểm (Mặc định)</th>
                        <th>Bí thư Chi bộ</th>
                        <th className="w-64">Ghi chú biến động đảng viên</th>
                        <th className="text-center">Trạng thái</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentMonthSchedules.map((item, idx) => (
                        <tr key={item.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs text-red-600 dark:text-red-400">{item.chiBo}</td>
                          <td className="text-center">
                            <input
                              type="number"
                              value={item.sl}
                              onChange={(e) => handleUpdateSchedule(item.chiBoId, 'sl', Number(e.target.value))}
                              className="w-16 text-center text-xs font-bold"
                            />
                          </td>
                          <td>
                            <div className="flex items-center gap-1.5">
                              <input
                                type="date"
                                min={`${selectedYear}-${String(selectedMonth).padStart(2, '0')}-01`}
                                max={`${selectedYear}-${String(selectedMonth).padStart(2, '0')}-${new Date(selectedYear, selectedMonth, 0).getDate()}`}
                                onChange={(e) => {
                                  if (e.target.value) {
                                    const [yyyy, mm, dd] = e.target.value.split('-');
                                    const d = new Date(e.target.value);
                                    const days = ['Chủ Nhật', 'thứ Hai', 'thứ Ba', 'thứ Tư', 'thứ Năm', 'thứ Sáu', 'thứ Bảy'];
                                    const dow = days[d.getDay()];
                                    const formatted = `${dd}/${mm}/${yyyy} (${dow})`;
                                    handleUpdateSchedule(item.chiBoId, 'thoiGian', formatted);
                                  }
                                }}
                                className="w-9 px-1 py-1 text-xs cursor-pointer bg-slate-50 dark:bg-slate-800"
                                title="Bấm vào để chọn ngày trên lịch (tự động hiện Thứ)"
                              />
                              <input
                                type="text"
                                placeholder="VD: 07/8/2026 (thứ Sáu)"
                                value={item.thoiGian === 'Chưa đăng ký' ? '' : item.thoiGian}
                                onChange={(e) => handleUpdateSchedule(item.chiBoId, 'thoiGian', e.target.value)}
                                className="flex-1 text-xs font-bold text-blue-700 dark:text-blue-300"
                              />
                            </div>
                          </td>
                          <td>
                            <input
                              type="text"
                              value={item.diaDiem}
                              onChange={(e) => handleUpdateSchedule(item.chiBoId, 'diaDiem', e.target.value)}
                              className="text-xs font-medium"
                            />
                          </td>
                          <td className="font-semibold text-xs">
                            <div>{item.biThu}</div>
                            <div className="text-[11px] font-mono text-[var(--text-muted)]">{item.sdt}</div>
                          </td>
                          <td>
                            <input
                              type="text"
                              placeholder="VD: Miễn sinh hoạt 01 đ/c..."
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
                          className="px-3 py-1.5 rounded bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 text-xs font-semibold"
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

          {/* TAB 3: ĐÁNH GIÁ XẾP LOẠI CHẤT LƯỢNG SINH HOẠT HẰNG THÁNG (ĐGXL) */}
          {activeTab === 'dgxl' && (
            <div className="space-y-6">
              <div className="card-glass p-5 flex flex-wrap items-center justify-between gap-4">
                <div>
                  <h3 className="text-base font-bold text-[var(--text-main)]">
                    Bảng Tổng Hợp Đánh Giá, Xếp Loại Chất Lượng Sinh Hoạt Chi Bộ Tháng {selectedMonth}/{selectedYear}
                  </h3>
                  <p className="text-xs text-[var(--text-muted)]">Theo dõi điểm số, mức xếp loại (Tốt, Khá, TB, Kém) và thuyết minh điểm trừ sau mỗi kỳ họp</p>
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
                    onClick={handleExportDGXLToExcel}
                    className="px-3 py-2 bg-emerald-600 text-white rounded-md text-xs font-bold flex items-center gap-1.5 shadow cursor-pointer"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" /> Xuất Excel ĐGXL
                  </button>
                </div>
              </div>

              <div className="card-glass p-1">
                <div className="table-container">
                  <table>
                    <thead>
                      <tr>
                        <th className="w-12 text-center">STT</th>
                        <th>Tên đơn vị</th>
                        <th className="w-20 text-center">Số lượng ĐV</th>
                        <th>Ngày họp chi bộ</th>
                        <th className="w-24 text-center">Điểm đánh giá</th>
                        <th className="w-36 text-center">Mức xếp loại</th>
                        <th>Ghi chú (Thuyết minh điểm trừ)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {currentMonthDGXL.map((d, idx) => (
                        <tr key={d.chiBoId}>
                          <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                          <td className="font-bold text-xs">{d.chiBo}</td>
                          <td className="text-center font-bold text-xs">{d.sl}</td>
                          <td className="text-xs font-medium text-blue-600 dark:text-blue-400">{d.ngayHop}</td>
                          <td className="text-center">
                            <input
                              type="number"
                              value={d.diemDG}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'diemDG', Number(e.target.value))}
                              className="w-16 text-center font-bold text-xs"
                            />
                          </td>
                          <td className="text-center">
                            <select
                              value={d.mucXepLoai}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'mucXepLoai', e.target.value)}
                              className="text-xs font-bold"
                            >
                              <option value="Tốt">Tốt</option>
                              <option value="Khá">Khá</option>
                              <option value="Trung bình">Trung bình</option>
                              <option value="Kém">Kém</option>
                            </select>
                          </td>
                          <td>
                            <input
                              type="text"
                              placeholder="Ghi rõ lý do trừ điểm..."
                              value={d.ghiChuTruDiem}
                              onChange={(e) => handleUpdateDGXL(d.chiBoId, 'ghiChuTruDiem', e.target.value)}
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

        {/* MODAL: TẠO THÔNG BÁO VÀ LẤY LINK */}
        {isShareModalOpen && (
          <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
            <div className="bg-[var(--bg-card)] border border-[var(--border-color)] rounded-2xl w-full max-w-xl overflow-hidden shadow-2xl animate-in fade-in zoom-in-95 duration-150">
              <div className="p-5 border-b border-[var(--border-color)] flex items-center justify-between bg-indigo-50/50 dark:bg-indigo-950/20">
                <div className="flex items-center gap-2">
                  <Share2 className="w-5 h-5 text-indigo-600" />
                  <h3 className="font-bold text-base text-[var(--text-main)]">
                    Tạo Thông Báo & Gửi Link Cho Các Đội Tự Đăng Ký
                  </h3>
                </div>
                <button onClick={() => setIsShareModalOpen(false)} className="text-lg font-bold cursor-pointer">✕</button>
              </div>

              <div className="p-6 space-y-4">
                <div>
                  <label className="block text-xs font-bold mb-1">Link trực tiếp gửi cho các Đội (Mở là chọn ngày được ngay):</label>
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
                      {copiedLink ? 'Đã sao chép' : 'Sao chép Link'}
                    </button>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold mb-1">Mẫu tin nhắn soạn sẵn gửi nhóm Zalo các Bí thư / Đội trưởng:</label>
                  <textarea
                    rows={6}
                    readOnly
                    value={sampleZaloMessage}
                    className="text-xs font-sans bg-slate-50 dark:bg-slate-900 p-3 leading-relaxed"
                  />
                  <button
                    onClick={() => {
                      navigator.clipboard.writeText(sampleZaloMessage);
                      setCopiedMessage(true);
                      setTimeout(() => setCopiedMessage(false), 2000);
                    }}
                    className="mt-2 w-full py-2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs rounded-md flex items-center justify-center gap-1 cursor-pointer"
                  >
                    {copiedMessage ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
                    {copiedMessage ? 'Đã sao chép tin nhắn Zalo' : 'Sao chép Toàn bộ Tin Nhắn gửi Zalo'}
                  </button>
                </div>

                <div className="p-3 rounded-lg bg-blue-50 dark:bg-blue-950/40 text-xs text-blue-800 dark:text-blue-200">
                  💡 <b>Sau khi up lên Vercel:</b> Link trên sẽ tự động đổi thành link online (ví dụ: <code>https://quanlydang-qltt.vercel.app/?mode=register</code>). Các Đội mở trên điện thoại hay máy tính đều tự chọn ngày được ngay!
                </div>
              </div>
            </div>
          </div>
        )}

      </div>
    </div>
  );
}
