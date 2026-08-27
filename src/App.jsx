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
  ExternalLink,
  Check
} from 'lucide-react';

export default function App() {
  const [darkMode, setDarkMode] = useState(false);

  // Check URL params for Member Registration Mode
  const urlParams = new URLSearchParams(window.location.search);
  const isPublicRegisterMode = urlParams.get('mode') === 'register';
  const paramChiBoId = urlParams.get('chibo');

  // Core Data States
  const [members, setMembers] = useState(() => {
    const saved = localStorage.getItem('qltt_party_members_v2');
    return saved ? JSON.parse(saved) : INITIAL_MEMBERS;
  });

  const [movements, setMovements] = useState(() => {
    const saved = localStorage.getItem('qltt_party_movements_v2');
    return saved ? JSON.parse(saved) : INITIAL_MOVEMENTS;
  });

  const [chiBoList] = useState(INITIAL_CHI_BO);
  const [branchDetails] = useState(INITIAL_BRANCH_DETAILS);

  const [evaluations, setEvaluations] = useState(() => {
    const saved = localStorage.getItem('qltt_party_evaluations_v2');
    return saved ? JSON.parse(saved) : INITIAL_EVALUATIONS;
  });

  const [fees, setFees] = useState(() => {
    const saved = localStorage.getItem('qltt_party_fees_v2');
    return saved ? JSON.parse(saved) : INITIAL_FEES;
  });

  // Lịch sinh hoạt 13 chi bộ theo tháng
  const [meetingSchedules, setMeetingSchedules] = useState(() => {
    const saved = localStorage.getItem('qltt_party_meetings');
    return saved ? JSON.parse(saved) : INITIAL_MEETING_SCHEDULES;
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
  const [teamSelectChiBo, setTeamSelectChiBo] = useState(paramChiBoId ? Number(paramChiBoId) : 1);
  const [teamDate, setTeamDate] = useState('');
  const [teamDayOfWeek, setTeamDayOfWeek] = useState('thứ Hai');
  const [teamLocation, setTeamLocation] = useState('');
  const [teamSubmitted, setTeamSubmitted] = useState(false);

  // Modal Member
  const [isMemberModalOpen, setIsMemberModalOpen] = useState(false);
  const [editingMember, setEditingMember] = useState(null);
  const [memberFormData, setMemberFormData] = useState({
    hoTen: '',
    gioiTinh: 'Nam',
    ngaySinh: '',
    soTheDang: '',
    soLyLich: '',
    chiBo: INITIAL_CHI_BO[0],
    chucVuDang: 'Đảng viên',
    chucVuChinhQuyen: 'Kiểm soát viên',
    trinhDoChuyenMon: 'Đại học',
    lyLuanChinhTri: 'Sơ cấp',
    ngayVaoDang: '',
    ngayChinhThuc: '',
    loaiDangVien: 'Chính thức',
    trangThai: 'Đang sinh hoạt',
    soDienThoai: '',
    email: '',
    queQuan: ''
  });

  useEffect(() => {
    localStorage.setItem('qltt_party_members_v2', JSON.stringify(members));
  }, [members]);

  useEffect(() => {
    localStorage.setItem('qltt_party_movements_v2', JSON.stringify(movements));
  }, [movements]);

  useEffect(() => {
    localStorage.setItem('qltt_party_evaluations_v2', JSON.stringify(evaluations));
  }, [evaluations]);

  useEffect(() => {
    localStorage.setItem('qltt_party_fees_v2', JSON.stringify(fees));
  }, [fees]);

  useEffect(() => {
    localStorage.setItem('qltt_party_meetings', JSON.stringify(meetingSchedules));
  }, [meetingSchedules]);

  const currentKey = `${selectedYear}-${selectedMonth}`;
  const currentMonthSchedules = useMemo(() => {
    if (meetingSchedules[currentKey]) {
      return meetingSchedules[currentKey];
    }
    return branchDetails.map(b => ({
      chiBoId: b.id,
      chiBo: b.name,
      sl: b.sl,
      thoiGian: 'Chưa đăng ký',
      diaDiem: b.diaDiem,
      biThu: b.biThu,
      sdt: b.sdt,
      trangThai: 'Chờ đăng ký'
    }));
  }, [meetingSchedules, currentKey, branchDetails]);

  // Handler cập nhật lịch họp của 1 Chi bộ
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

  // Duyệt thống nhất toàn bộ lịch họp với Lãnh đạo
  const handleApproveAllSchedules = () => {
    const updatedList = currentMonthSchedules.map(item => ({
      ...item,
      trangThai: 'Đã thống nhất'
    }));
    setMeetingSchedules(prev => ({
      ...prev,
      [currentKey]: updatedList
    }));
    alert(`Đã duyệt thống nhất lịch sinh hoạt tháng ${selectedMonth}/${selectedYear} với Lãnh đạo! Đồng chí có thể bấm 'In Thông Báo Trình Ký'.`);
  };

  // Đội tự gửi đăng ký
  const handleTeamSubmitRegistration = (e) => {
    e.preventDefault();
    if (!teamDate) {
      alert('Vui lòng chọn ngày họp!');
      return;
    }
    // Format date string: DD/MM/YYYY (thứ ...)
    const [yyyy, mm, dd] = teamDate.split('-');
    const timeFormatted = `${dd}/${mm}/${yyyy} (${teamDayOfWeek})`;
    
    const targetBranch = branchDetails.find(b => b.id === teamSelectChiBo);
    const loc = teamLocation || (targetBranch ? targetBranch.diaDiem : 'Tại đơn vị');

    handleUpdateSchedule(teamSelectChiBo, 'thoiGian', timeFormatted);
    handleUpdateSchedule(teamSelectChiBo, 'diaDiem', loc);

    setTeamSubmitted(true);
    setTimeout(() => setTeamSubmitted(false), 5000);
  };

  const handlePrint = () => {
    window.print();
  };

  // Text thông báo gửi Zalo / Văn bản cho các Đội
  const sampleZaloMessage = `[THÔNG BÁO ĐẢNG ỦY BỘ PHẬN CHI CỤC QLTT]
Kính gửi: Bí thư các Chi bộ trực thuộc (Đội 1 đến Đội 12 và Khối phòng).
Thực hiện Quy chế làm việc, đề nghị các Chi bộ chủ động đăng ký lịch sinh hoạt lệ tháng ${selectedMonth}/${selectedYear} trước ngày 20 để Đảng ủy tổng hợp xin ý kiến Lãnh đạo và ban hành Thông báo chính thức.
👉 Link đăng ký trực tuyến: ${window.location.origin + window.location.pathname}?mode=register`;

  // GIAO DIỆN DÀNH RIÊNG CHO CÁC ĐỘI TỰ ĐĂNG KÝ
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
              <div className="p-4 rounded-xl bg-emerald-50 border border-emerald-200 text-center space-y-2">
                <CheckCircle2 className="w-10 h-10 text-emerald-600 mx-auto" />
                <h4 className="font-bold text-emerald-800">Đăng ký thành công!</h4>
                <p className="text-xs text-emerald-700">
                  Lịch sinh hoạt của <b>{selectedBranchInfo.name}</b> đã được gửi về Đảng ủy bộ phận để tổng hợp trình Lãnh đạo.
                </p>
                <button
                  type="button"
                  onClick={() => setTeamSubmitted(false)}
                  className="mt-2 text-xs text-blue-600 font-bold underline"
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
                    onChange={(e) => {
                      const cid = Number(e.target.value);
                      setTeamSelectChiBo(cid);
                      const br = branchDetails.find(b => b.id === cid);
                      if (br) setTeamLocation(br.diaDiem);
                    }}
                    className="w-full p-2.5 border rounded-lg text-sm font-semibold bg-slate-50"
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
                  <label className="block text-xs font-bold text-slate-700 mb-1">3. Địa điểm tổ chức sinh hoạt</label>
                  <input
                    type="text"
                    value={teamLocation}
                    onChange={(e) => setTeamLocation(e.target.value)}
                    placeholder="VD: Đội QLTT số 2 hoặc Hội trường..."
                    className="w-full p-2.5 border rounded-lg text-sm bg-white"
                  />
                </div>

                <div className="p-3 rounded-lg bg-amber-50 border border-amber-200 text-xs text-amber-800">
                  📌 <b>Lưu ý:</b> Đề nghị các Chi bộ hoàn thành đăng ký trước ngày 20 hàng tháng để cán bộ tổng hợp báo cáo Đảng ủy.
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

  // GIAO DIỆN CHÍNH (CÁN BỘ TỔNG HỢP)
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
                    ĐẢNG ỦY BỘ PHẬN CHI CỤC QUẢN LÝ THỊ TRƯỜNG
                  </h1>
                  <span className="badge bg-red-50 text-red-700 border border-red-200 dark:bg-red-950/40 dark:text-red-300 dark:border-red-800">
                    Tổng hợp 13 Chi bộ
                  </span>
                </div>
                <p className="text-xs text-[var(--text-muted)]">Hệ thống tạo thông báo, gửi link đăng ký ngày 20 & phê duyệt trình ký</p>
              </div>
            </div>

            <div className="flex items-center gap-2.5">
              {/* NÚT TẠO THÔNG BÁO & GỬI LINK CHO ĐỘI */}
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
                title="In thông báo lịch sinh hoạt trình ký Bí thư Đảng ủy"
              >
                <Printer className="w-4 h-4" />
                In Thông Báo Trình Ký
              </button>

              <button 
                onClick={() => setDarkMode(!darkMode)}
                className="p-2 rounded-lg border border-[var(--border-color)] text-[var(--text-muted)] hover:text-[var(--text-main)] transition-all cursor-pointer"
              >
                {darkMode ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {/* Navigation Tabs */}
          <div className="max-w-7xl mx-auto flex gap-6 mt-3 border-t border-[var(--border-color)] pt-2 overflow-x-auto">
            {[
              { id: 'meetings', label: '📅 Lịch Sinh Hoạt & Đăng Ký (Ngày 20)', icon: Calendar },
              { id: 'dashboard', label: '📊 Tổng quan Đảng bộ (157 ĐV)', icon: Sparkles },
              { id: 'members', label: '👥 Danh sách Đảng viên', icon: Users },
              { id: 'movements', label: '📈 Biến động Tăng/Giảm', icon: TrendingUp },
              { id: 'fees', label: '💳 Đóng Đảng phí 12 Tháng', icon: CreditCard },
              { id: 'evaluations', label: '🏆 Đánh giá & Xếp loại', icon: Award }
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

        {/* PRINT DOCUMENT FORMAT */}
        <div className="hidden print:block p-10 text-black bg-white" style={{ fontFamily: 'Times New Roman, serif' }}>
          <div className="flex justify-between items-start text-center mb-6">
            <div className="w-1/2">
              <p className="text-sm font-normal">ĐẢNG BỘ SỞ CÔNG THƯƠNG TỈNH AN GIANG</p>
              <p className="text-sm font-bold uppercase underline">ĐẢNG ỦY BỘ PHẬN CHI CỤC QUẢN LÝ THỊ TRƯỜNG</p>
              <p className="text-xs mt-1">*</p>
              <p className="text-xs">Số: &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; -TB/ĐU</p>
            </div>
            <div className="w-1/2">
              <p className="text-sm font-bold">ĐẢNG CỘNG SẢN VIỆT NAM</p>
              <p className="text-xs italic mt-2">An Giang, ngày &nbsp;&nbsp;&nbsp;&nbsp; tháng {selectedMonth} năm {selectedYear}</p>
            </div>
          </div>

          <div className="text-center my-6">
            <h2 className="text-lg font-bold uppercase">THÔNG BÁO</h2>
            <h3 className="text-base font-bold">Lịch sinh hoạt lệ tháng {selectedMonth}/{selectedYear} của các chi bộ trực thuộc</h3>
            <p className="text-xs">-----</p>
          </div>

          <p className="text-sm leading-relaxed mb-2" style={{ textIndent: '30px' }}>
            Căn cứ Quy chế làm việc của Đảng ủy Chi cục Quản lý thị trường và Quy chế làm việc của các chi bộ trực thuộc nhiệm kỳ 2025-2030.
          </p>
          <p className="text-sm leading-relaxed mb-4" style={{ textIndent: '30px' }}>
            Theo đăng ký lịch sinh hoạt lệ chi bộ tháng {String(selectedMonth).padStart(2, '0')}/{selectedYear}. Đảng ủy bộ phận Chi cục Quản lý thị trường thông báo thời gian, địa điểm sinh hoạt của các chi bộ, như sau:
          </p>

          <table className="w-full border-collapse border border-black text-sm my-4">
            <thead>
              <tr className="bg-gray-100 font-bold text-center">
                <th className="border border-black p-2 w-12">STT</th>
                <th className="border border-black p-2">Chi bộ</th>
                <th className="border border-black p-2 w-16">SL</th>
                <th className="border border-black p-2">Thời gian</th>
                <th className="border border-black p-2">Địa điểm</th>
                <th className="border border-black p-2">Bí thư, điện thoại</th>
              </tr>
            </thead>
            <tbody>
              {currentMonthSchedules.map((item, idx) => (
                <tr key={idx}>
                  <td className="border border-black p-2 text-center">{idx + 1}</td>
                  <td className="border border-black p-2 font-bold">{item.chiBo}</td>
                  <td className="border border-black p-2 text-center">{item.sl}</td>
                  <td className="border border-black p-2 text-center">{item.thoiGian}</td>
                  <td className="border border-black p-2">{item.diaDiem}</td>
                  <td className="border border-black p-2">
                    <div className="font-semibold">{item.biThu}</div>
                    <div className="text-xs font-mono">{item.sdt}</div>
                  </td>
                </tr>
              ))}
              <tr className="font-bold text-center bg-gray-50">
                <td colSpan="2" className="border border-black p-2">TỔNG SỐ</td>
                <td className="border border-black p-2 text-center">{currentMonthSchedules.reduce((sum, i) => sum + (Number(i.sl) || 0), 0)}</td>
                <td colSpan="3" className="border border-black p-2"></td>
              </tr>
            </tbody>
          </table>

          <div className="flex justify-between items-start mt-8 text-sm">
            <div className="w-1/2 italic text-xs">
              <p className="font-bold underline not-italic">Nơi nhận:</p>
              <p>- Đảng ủy Sở Công Thương;</p>
              <p>- Bí thư các Chi bộ trực thuộc;</p>
              <p>- Lưu: Đảng ủy.</p>
            </div>
            <div className="w-1/2 text-center">
              <p className="font-bold uppercase">T/M ĐẢNG ỦY</p>
              <p className="font-bold uppercase">BÍ THƯ</p>
              <div className="h-24"></div>
              <p className="font-bold uppercase">Nguyễn Trung Tiến</p>
            </div>
          </div>
        </div>

        {/* MAIN BODY */}
        <main className="max-w-7xl mx-auto p-6 flex-1 w-full">

          {/* TAB: LỊCH HỌP */}
          {activeTab === 'meetings' && (
            <div className="space-y-6">
              
              {/* QUY TRÌNH HÀNG THÁNG BANNER */}
              <div className="card-glass p-5 border-l-4 border-amber-500 bg-linear-to-r from-amber-500/5 to-transparent">
                <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                  <div>
                    <div className="flex items-center gap-2">
                      <Clock className="w-5 h-5 text-amber-600" />
                      <h3 className="text-base font-bold text-[var(--text-main)]">
                        Quy trình Quản lý Lịch họp Chi bộ định kỳ (Ngày 20 hàng tháng)
                      </h3>
                    </div>
                    <p className="text-xs text-[var(--text-muted)] mt-1">
                      1. Bấm <b>'Gửi Link Đăng Ký'</b> gửi Zalo các Đội $ightarrow$ 2. Đội tự chọn ngày $ightarrow$ 3. Trình Lãnh đạo thống nhất $ightarrow$ 4. Bấm In trình ký Bí thư.
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
                    Tổng số đảng viên: {currentMonthSchedules.reduce((sum, i) => sum + (Number(i.sl) || 0), 0)} ĐV
                  </span>
                </div>

                <div className="flex items-center gap-2.5">
                  <button
                    onClick={() => setIsShareModalOpen(true)}
                    className="px-3.5 py-2 rounded-md bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                  >
                    <Share2 className="w-4 h-4" />
                    Lấy Link Gửi Các Đội
                  </button>

                  <button
                    onClick={handleApproveAllSchedules}
                    className="px-3.5 py-2 rounded-md bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                    title="Đánh dấu lãnh đạo đã đồng ý thông nhất toàn bộ"
                  >
                    <CheckCheck className="w-4 h-4" />
                    Lãnh đạo đã Thống nhất
                  </button>

                  <button
                    onClick={handlePrint}
                    className="px-3.5 py-2 rounded-md bg-red-600 hover:bg-red-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                  >
                    <Printer className="w-4 h-4" />
                    In Thông Báo (Trình Ký)
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
                        <th className="w-64">Thời gian sinh hoạt (Đăng ký)</th>
                        <th>Địa điểm</th>
                        <th>Bí thư Chi bộ</th>
                        <th>Số điện thoại</th>
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
                            <input
                              type="text"
                              placeholder="VD: 07/8/2026 (thứ Sáu)"
                              value={item.thoiGian === 'Chưa đăng ký' ? '' : item.thoiGian}
                              onChange={(e) => handleUpdateSchedule(item.chiBoId, 'thoiGian', e.target.value)}
                              className="text-xs font-semibold"
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
                          <td className="font-semibold text-xs">{item.biThu}</td>
                          <td className="font-mono text-xs text-[var(--text-muted)]">{item.sdt}</td>
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

          {/* TAB: DASHBOARD */}
          {activeTab === 'dashboard' && (
            <div className="space-y-6">
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <div className="card-glass p-5 flex items-center justify-between border-l-4 border-red-500">
                  <div>
                    <p className="text-xs font-medium text-[var(--text-muted)] uppercase">Tổng số Đảng viên</p>
                    <h3 className="text-3xl font-extrabold mt-1 text-[var(--text-main)]">{totalCount}</h3>
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
                    <h3 className="text-3xl font-extrabold mt-1 text-[var(--text-main)]">{officialCount}</h3>
                    <p className="text-xs text-[var(--text-muted)] mt-1">Tỷ lệ: 97%</p>
                  </div>
                  <div className="w-12 h-12 rounded-xl bg-blue-50 dark:bg-blue-950/50 flex items-center justify-center text-blue-600">
                    <ShieldCheck className="w-6 h-6" />
                  </div>
                </div>

                <div className="card-glass p-5 flex items-center justify-between border-l-4 border-amber-500">
                  <div>
                    <p className="text-xs font-medium text-[var(--text-muted)] uppercase">Dự bị</p>
                    <h3 className="text-3xl font-extrabold mt-1 text-amber-600">{probationaryCount}</h3>
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
                    <h3 className="text-3xl font-extrabold mt-1 text-[var(--text-main)]">{femaleCount}</h3>
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

          {/* TAB: MEMBERS */}
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

                <button
                  onClick={() => setIsMemberModalOpen(true)}
                  className="px-4 py-2 rounded-md bg-red-600 hover:bg-red-700 text-white font-semibold text-xs flex items-center gap-1.5 shadow-sm cursor-pointer"
                >
                  <UserPlus className="w-4 h-4" />
                  Thêm Đảng viên
                </button>
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
                      {filteredMembers.map((m, idx) => (
                        <tr key={m.id}>
                          <td className="text-center text-[var(--text-muted)] text-xs">{idx + 1}</td>
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

          {/* TAB: MOVEMENTS */}
          {activeTab === 'movements' && (
            <div className="card-glass p-5 space-y-4">
              <h3 className="text-base font-bold text-[var(--text-main)]">Biến động Tăng/Giảm trong năm</h3>
              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      <th>STT</th>
                      <th>Thời gian</th>
                      <th>Họ và Tên</th>
                      <th>Chi bộ</th>
                      <th>Loại biến động</th>
                      <th>Số Quyết định</th>
                      <th>Ghi chú</th>
                    </tr>
                  </thead>
                  <tbody>
                    {movements.map((mv, idx) => (
                      <tr key={mv.id}>
                        <td className="text-center text-xs text-[var(--text-muted)]">{idx + 1}</td>
                        <td className="font-bold text-xs">Tháng {mv.thang}/{mv.nam}</td>
                        <td className="font-semibold">{mv.hoTen}</td>
                        <td className="text-xs">{mv.chiBo}</td>
                        <td>
                          <span className="badge bg-emerald-50 text-emerald-700 border border-emerald-200">
                            {mv.loaiBienDong}
                          </span>
                        </td>
                        <td className="text-xs font-mono">{mv.soQuyetDinh}</td>
                        <td className="text-xs text-[var(--text-muted)]">{mv.ghiChu}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB: FEES */}
          {activeTab === 'fees' && (
            <div className="card-glass p-5 space-y-4">
              <h3 className="text-base font-bold text-[var(--text-main)]">Bảng theo dõi Nộp Đảng phí 12 Tháng</h3>
              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      <th>STT</th>
                      <th>Họ và Tên</th>
                      <th>Chi bộ</th>
                      {[1,2,3,4,5,6,7,8,9,10,11,12].map(m => <th key={m} className="text-center">T{m}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {members.map((m, idx) => (
                      <tr key={m.id}>
                        <td className="text-center text-xs">{idx + 1}</td>
                        <td className="font-bold text-xs">{m.hoTen}</td>
                        <td className="text-xs">{m.chiBo}</td>
                        {[1,2,3,4,5,6,7,8,9,10,11,12].map(mon => (
                          <td key={mon} className="text-center text-xs font-bold text-emerald-600">✓</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB: EVALUATIONS */}
          {activeTab === 'evaluations' && (
            <div className="card-glass p-5 space-y-4">
              <h3 className="text-base font-bold text-[var(--text-main)]">Đánh giá & Xếp loại Cuối năm</h3>
              <div className="table-container">
                <table>
                  <thead>
                    <tr>
                      <th>STT</th>
                      <th>Họ và Tên</th>
                      <th>Chi bộ</th>
                      <th>Xếp loại chất lượng</th>
                      <th>Khen thưởng</th>
                      <th>Kỷ luật</th>
                    </tr>
                  </thead>
                  <tbody>
                    {members.map((m, idx) => (
                      <tr key={m.id}>
                        <td className="text-center text-xs">{idx + 1}</td>
                        <td className="font-bold text-xs">{m.hoTen}</td>
                        <td className="text-xs">{m.chiBo}</td>
                        <td className="text-xs font-semibold text-blue-600">Hoàn thành tốt nhiệm vụ</td>
                        <td className="text-xs">Giấy khen Cục</td>
                        <td className="text-xs text-emerald-600">Không</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

        </main>

        {/* MODAL: TẠO THÔNG BÁO VÀ LẤY LINK GỬI CHO CÁC ĐỘI */}
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
                <button onClick={() => setIsShareModalOpen(false)} className="text-lg font-bold">✕</button>
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
