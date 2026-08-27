// Dữ liệu chuẩn Đảng bộ bộ phận Chi cục Quản lý thị trường (13 Chi bộ)
export const INITIAL_CHI_BO = [
  'Chi bộ 1',
  'Chi bộ 2',
  'Chi bộ 3',
  'Chi bộ 4',
  'Chi bộ 5',
  'Chi bộ 6',
  'Chi bộ 7',
  'Chi bộ 8',
  'Chi bộ 9',
  'Chi bộ 10',
  'Chi bộ 11',
  'Chi bộ 12',
  'Chi bộ Khối phòng'
];

export const INITIAL_BRANCH_DETAILS = [
  { id: 1, name: 'Chi bộ 1', diaDiem: 'VP Chi cục', sl: 18, biThu: 'Đặng Thanh Phê', sdt: '0918.233.352' },
  { id: 2, name: 'Chi bộ 2', diaDiem: 'Đội QLTT số 2', sl: 14, biThu: 'Đào Minh Phúc', sdt: '0943.685.678' },
  { id: 3, name: 'Chi bộ 3', diaDiem: 'Đội QLTT số 3', sl: 8, biThu: 'Nguyễn Hữu Thọ', sdt: '0915.755.761' },
  { id: 4, name: 'Chi bộ 4', diaDiem: 'Đội QLTT số 4', sl: 8, biThu: 'Dương Thành Sự', sdt: '0919.999.182' },
  { id: 5, name: 'Chi bộ 5', diaDiem: 'Đội QLTT số 5', sl: 10, biThu: 'Trương Cáo', sdt: '0939.949.499' },
  { id: 6, name: 'Chi bộ 6', diaDiem: 'Đội QLTT số 6', sl: 9, biThu: 'Bùi Phước Lan', sdt: '0919.922.773' },
  { id: 7, name: 'Chi bộ 7', diaDiem: 'Đội QLTT số 7', sl: 8, biThu: 'Ngô Chí Trung', sdt: '0911.658.288' },
  { id: 8, name: 'Chi bộ 8', diaDiem: 'Đội QLTT số 8', sl: 10, biThu: 'Diệp Trọng Danh', sdt: '0913.125.981' },
  { id: 9, name: 'Chi bộ 9', diaDiem: 'Đội QLTT số 9', sl: 13, biThu: 'Trần Thị Thu Thanh Thủy', sdt: '0989.625.878' },
  { id: 10, name: 'Chi bộ 10', diaDiem: 'Đội QLTT số 10', sl: 12, biThu: 'Nguyễn Phúc Xuân Thụy', sdt: '0918.823.001' },
  { id: 11, name: 'Chi bộ 11', diaDiem: 'Đội QLTT số 11', sl: 12, biThu: 'Phan Thành Phục', sdt: '0907.776.852' },
  { id: 12, name: 'Chi bộ 12', diaDiem: 'Đội QLTT số 12', sl: 16, biThu: 'Lê Trọng Hiếu', sdt: '0944.844.444' },
  { id: 13, name: 'Chi bộ Khối phòng', diaDiem: 'Hội trường B', sl: 19, biThu: 'Trần Quang Thái', sdt: '0919.999.101' }
];

export const INITIAL_MEMBERS = [
  {
    id: 'DV001',
    hoTen: 'Đặng Thanh Phê',
    gioiTinh: 'Nam',
    ngaySinh: '1975-06-15',
    soTheDang: '18023456',
    soLyLich: 'LL-001',
    chiBo: 'Chi bộ 1',
    chucVuDang: 'Bí thư Chi bộ',
    chucVuChinhQuyen: 'Đội trưởng',
    trinhDoChuyenMon: 'Đại học',
    lyLuanChinhTri: 'Cao cấp',
    ngayVaoDang: '2004-02-03',
    ngayChinhThuc: '2005-02-03',
    loaiDangVien: 'Chính thức',
    trangThai: 'Đang sinh hoạt',
    soDienThoai: '0918.233.352',
    email: '',
    queQuan: 'An Giang'
  },
  {
    id: 'DV002',
    hoTen: 'Đào Minh Phúc',
    gioiTinh: 'Nam',
    ngaySinh: '1979-08-20',
    soTheDang: '18045678',
    soLyLich: 'LL-002',
    chiBo: 'Chi bộ 2',
    chucVuDang: 'Bí thư Chi bộ',
    chucVuChinhQuyen: 'Đội trưởng',
    trinhDoChuyenMon: 'Cử nhân Kinh tế',
    lyLuanChinhTri: 'Trung cấp',
    ngayVaoDang: '2010-05-19',
    ngayChinhThuc: '2011-05-19',
    loaiDangVien: 'Chính thức',
    trangThai: 'Đang sinh hoạt',
    soDienThoai: '0943.685.678',
    email: '',
    queQuan: 'An Giang'
  },
  {
    id: 'DV003',
    hoTen: 'Trần Thị Thu Thanh Thủy',
    gioiTinh: 'Nữ',
    ngaySinh: '1984-11-15',
    soTheDang: '18067890',
    soLyLich: 'LL-009',
    chiBo: 'Chi bộ 9',
    chucVuDang: 'Bí thư Chi bộ',
    chucVuChinhQuyen: 'Đội trưởng',
    trinhDoChuyenMon: 'Cử nhân Luật',
    lyLuanChinhTri: 'Trung cấp',
    ngayVaoDang: '2015-09-02',
    ngayChinhThuc: '2016-09-02',
    loaiDangVien: 'Chính thức',
    trangThai: 'Đang sinh hoạt',
    soDienThoai: '0989.625.878',
    email: '',
    queQuan: 'An Giang'
  },
  {
    id: 'DV004',
    hoTen: 'Nguyễn Văn Dự Bị',
    gioiTinh: 'Nam',
    ngaySinh: '1998-03-25',
    soTheDang: 'Chưa cấp',
    soLyLich: 'LL-015',
    chiBo: 'Chi bộ Khối phòng',
    chucVuDang: 'Đảng viên',
    chucVuChinhQuyen: 'Kiểm soát viên',
    trinhDoChuyenMon: 'Đại học Thương mại',
    lyLuanChinhTri: 'Sơ cấp',
    ngayVaoDang: '2025-09-15',
    ngayChinhThuc: '',
    loaiDangVien: 'Dự bị',
    trangThai: 'Đang sinh hoạt',
    soDienThoai: '0912345678',
    email: '',
    queQuan: 'An Giang'
  }
];

export const INITIAL_MOVEMENTS = [
  {
    id: 'BD001',
    nam: 2026,
    thang: 2,
    dangVienId: 'DV004',
    hoTen: 'Nguyễn Văn Dự Bị',
    chiBo: 'Chi bộ Khối phòng',
    loaiBienDong: 'Kết nạp mới',
    ngayThucHien: '2026-02-03',
    soQuyetDinh: 'QĐ-08/ĐU-QLTT',
    ghiChu: 'Kết nạp đợt 3/2'
  }
];

export const INITIAL_EVALUATIONS = {
  '2025': {
    'DV001': { xepLoai: 'Hoàn thành xuất sắc nhiệm vụ', khenThuong: 'Chiến sĩ thi đua cấp cơ sở', kyLuat: 'Không' },
    'DV002': { xepLoai: 'Hoàn thành tốt nhiệm vụ', khenThuong: 'Giấy khen Cục QLTT', kyLuat: 'Không' }
  }
};

export const INITIAL_FEES = {
  '2026': {
    'DV001': { 1: true, 2: true, 3: true, 4: true, 5: true, 6: true, 7: true, 8: true, 9: false, 10: false, 11: false, 12: false },
    'DV002': { 1: true, 2: true, 3: true, 4: true, 5: true, 6: true, 7: true, 8: true, 9: false, 10: false, 11: false, 12: false }
  }
};

// Lịch sinh hoạt chi bộ hàng tháng (13 chi bộ)
export const INITIAL_MEETING_SCHEDULES = {
  '2026-8': [
    { chiBoId: 1, chiBo: 'Chi bộ 1', sl: 18, thoiGian: '07/8/2026 (thứ Sáu)', diaDiem: 'VP Chi cục', biThu: 'Đặng Thanh Phê', sdt: '0918.233.352', trangThai: 'Đã thống nhất' },
    { chiBoId: 2, chiBo: 'Chi bộ 2', sl: 14, thoiGian: '03/8/2026 (thứ Hai)', diaDiem: 'Đội QLTT số 2', biThu: 'Đào Minh Phúc', sdt: '0943.685.678', trangThai: 'Đã thống nhất' },
    { chiBoId: 3, chiBo: 'Chi bộ 3', sl: 8, thoiGian: '03/8/2026 (thứ Hai)', diaDiem: 'Đội QLTT số 3', biThu: 'Nguyễn Hữu Thọ', sdt: '0915.755.761', trangThai: 'Đã thống nhất' },
    { chiBoId: 4, chiBo: 'Chi bộ 4', sl: 8, thoiGian: '07/8/2026 (Thứ Sáu)', diaDiem: 'Đội QLTT số 4', biThu: 'Dương Thành Sự', sdt: '0919.999.182', trangThai: 'Đã thống nhất' },
    { chiBoId: 5, chiBo: 'Chi bộ 5', sl: 10, thoiGian: '05/8/2026 (thứ Tư)', diaDiem: 'Đội QLTT số 5', biThu: 'Trương Cáo', sdt: '0939.949.499', trangThai: 'Đã thống nhất' },
    { chiBoId: 6, chiBo: 'Chi bộ 6', sl: 9, thoiGian: '07/8/2026 (thứ Sáu)', diaDiem: 'Đội QLTT số 6', biThu: 'Bùi Phước Lan', sdt: '0919.922.773', trangThai: 'Đã thống nhất' },
    { chiBoId: 7, chiBo: 'Chi bộ 7', sl: 8, thoiGian: '03/8/2026 (thứ Hai)', diaDiem: 'Đội QLTT số 7', biThu: 'Ngô Chí Trung', sdt: '0911.658.288', trangThai: 'Đã thống nhất' },
    { chiBoId: 8, chiBo: 'Chi bộ 8', sl: 10, thoiGian: '03/8/2026 (thứ Hai)', diaDiem: 'Đội QLTT số 8', biThu: 'Diệp Trọng Danh', sdt: '0913.125.981', trangThai: 'Đã thống nhất' },
    { chiBoId: 9, chiBo: 'Chi bộ 9', sl: 13, thoiGian: '07/8/2026 (thứ Sáu)', diaDiem: 'Đội QLTT số 9', biThu: 'Trần Thị Thu Thanh Thủy', sdt: '0989.625.878', trangThai: 'Đã thống nhất' },
    { chiBoId: 10, chiBo: 'Chi bộ 10', sl: 12, thoiGian: '05/8/2026 (thứ Tư)', diaDiem: 'Đội QLTT số 10', biThu: 'Nguyễn Phúc Xuân Thụy', sdt: '0918.823.001', trangThai: 'Đã thống nhất' },
    { chiBoId: 11, chiBo: 'Chi bộ 11', sl: 12, thoiGian: '04/8/2026 (thứ Ba)', diaDiem: 'Đội QLTT số 11', biThu: 'Phan Thành Phục', sdt: '0907.776.852', trangThai: 'Đã thống nhất' },
    { chiBoId: 12, chiBo: 'Chi bộ 12', sl: 16, thoiGian: '05/8/2026 (thứ Tư)', diaDiem: 'Đội QLTT số 12', biThu: 'Lê Trọng Hiếu', sdt: '0944.844.444', trangThai: 'Đã thống nhất' },
    { chiBoId: 13, chiBo: 'Chi bộ Khối phòng', sl: 19, thoiGian: '03/8/2026 (thứ Hai)', diaDiem: 'Hội trường B', biThu: 'Trần Quang Thái', sdt: '0919.999.101', trangThai: 'Đã thống nhất' }
  ]
};
