import 'dart:async';
import 'dart:convert';
import 'dart:io';
import 'dart:math' as math;
import 'dart:ui' show FontFeature;

import 'package:excel/excel.dart' as xl;
import 'package:flutter/material.dart' hide Text;
import 'package:flutter/material.dart' as mat show Text;
import 'package:flutter/services.dart' show rootBundle;
import 'package:intl/intl.dart';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';
import 'package:pdf/pdf.dart' as pdf;
import 'package:pdf/widgets.dart' as pw;
import 'package:share_plus/share_plus.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:sqflite/sqflite.dart';

Future<void> main() async {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const SalaryApp());
}

// ════════════════════════════════════════════════════════════════════════════
//  Hằng số & tiện ích
// ════════════════════════════════════════════════════════════════════════════

// ── Đa ngôn ngữ (Tiếng Việt / English) ────────────────────────────────────
// Chuỗi gốc viết bằng tiếng Việt. Khi chọn English, Text và tr() tra bảng dịch
// bên dưới (khớp nguyên câu hoặc theo mẫu cho câu có số liệu).

String currentLang = 'vi';
const supportedLangs = {'vi': 'Tiếng Việt', 'en': 'English'};

List<String> get _weekdayShort => currentLang == 'en'
    ? const ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    : const ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];

List<String> get _weekdayFull => currentLang == 'en'
    ? const ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    : const ['Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy', 'Chủ nhật'];

/// Tên loại ngày để hiển thị (giá trị lưu trữ vẫn là 'Thường' / 'OT' / 'Lễ').
String dayLabel(String type) {
  if (currentLang != 'en') return type;
  return switch (type) {
    'Lễ' => 'Holiday',
    'OT' => 'Weekend',
    _ => 'Weekday',
  };
}

const _enExact = <String, String>{
  'Theo dõi lương': 'Pay Tracker',
  'Xuất dữ liệu': 'Export data',
  'Cài đặt': 'Settings',
  'Thêm ca': 'Add shift',
  'Thêm ca làm': 'Add shift',
  'Sửa ca làm': 'Edit shift',
  'Bắt đầu ca làm': 'Start a shift',
  'Hôm nay': 'Today',
  'Tuần này': 'This week',
  'Tháng này': 'This month',
  'Tháng trước': 'Last month',
  'Năm nay': 'This year',
  'Tất cả': 'All',
  'Lịch sử ca làm': 'Shift history',
  'Thu nhập tuần này': 'This week’s earnings',
  'Vào ca': 'Start shift',
  'Ra ca': 'End shift',
  'Hủy ca': 'Cancel shift',
  'ĐANG TRONG CA': 'ON SHIFT',
  'Hủy': 'Cancel',
  'Lưu': 'Save',
  'Sửa': 'Edit',
  'Xóa': 'Delete',
  'Thử lại': 'Retry',
  'Đóng': 'Close',
  'Sao chép': 'Copy',
  'Xuất CSV': 'Export CSV',
  'Đã sao chép CSV.': 'CSV copied.',
  'Xuất Excel (.xlsx) và PDF chỉ có trong bản app đầy đủ (main.dart).':
      'Excel (.xlsx) and PDF export is only in the full app (main.dart).',
  'Lương theo giờ': 'Hourly rate',
  'Lương mặc định (mỗi giờ)': 'Default hourly rate',
  'Mức lương không hợp lệ.': 'Invalid hourly rate.',
  'Giờ ra phải khác giờ vào.': 'End time must differ from start time.',
  'Giờ ra phải sau giờ vào.': 'End time must be after start time.',
  'Một ca tối đa 24 giờ.': 'A shift can be at most 24 hours.',
  'Ca này trùng giờ với một ca đã có.': 'This shift overlaps an existing shift.',
  'Ca làm cần ít nhất 1 phút. Bấm "Hủy ca" nếu bạn bấm nhầm.':
      'A shift must last at least 1 minute. Tap "Cancel shift" if you tapped by mistake.',
  'Hủy ca đang làm?': 'Cancel the current shift?',
  'Ca này sẽ không được lưu.': 'This shift will not be saved.',
  'Xóa ca làm?': 'Delete this shift?',
  'Đã lưu.': 'Saved.',
  'Đã lưu ca làm.': 'Shift saved.',
  'Đã lưu. Ca dài bất thường, hãy kiểm tra lại giờ ra (chạm vào ca để sửa).':
      'Saved. This shift is unusually long. Check the end time (tap the shift to edit).',
  'Đã xóa.': 'Deleted.',
  'Chưa có dữ liệu để xuất.': 'There is no data to export.',
  'Không có ca làm nào trong khoảng này.': 'No shifts in this period.',
  'Chưa có ca làm nào': 'No shifts yet',
  'Bấm “Vào ca” để chấm công hoặc thêm một ca thủ công.':
      'Tap “Start shift” to clock in, or add a shift manually.',
  'Thử lại.': 'Retry.',
  'Khoảng thời gian': 'Period',
  'Định dạng': 'Format',
  'Chi tiết từng ca và tổng hợp theo tuần': 'Each shift, plus weekly totals',
  'Báo cáo để in hoặc gửi (chạy offline)': 'A report to print or send (works offline)',
  'Văn bản thuần, mở được bằng Excel hoặc Google Sheets':
      'Plain text; opens in Excel or Google Sheets',
  'Ngày lễ (tính x2)': 'Holidays (paid x2)',
  'Chưa đánh dấu ngày lễ nào.': 'No holidays marked.',
  'Thêm ngày lễ': 'Add holiday',
  'Chọn ngày lễ': 'Select a holiday',
  'Đã đủ 40 giờ thường. Giờ làm thêm tính OT.':
      'Full 40 regular hours reached. Extra hours count as OT.',
  'Giờ thường trong tuần': 'Regular hours this week',
  'Thường': 'Regular',
  'Tổng giờ': 'Total hours',
  'Giờ thường': 'Regular hours',
  'Tổng tiền': 'Total pay',
  'Tiền': 'Pay',
  'Số ca': 'Shifts',
  'Tuần': 'Week',
  'Theo tuần': 'By week',
  'Chi tiết ca làm': 'Shift details',
  'Ngày': 'Date',
  'Thứ': 'Day',
  'Vào': 'In',
  'Ra': 'Out',
  'Loại': 'Type',
  'Giờ vào': 'Start',
  'Giờ ra': 'End',
  'Loại ngày': 'Day type',
  'Lương/giờ': 'Rate/hr',
  'Lương/giờ (USD)': 'Rate/hr (USD)',
  'Tiền (USD)': 'Pay (USD)',
  'Giờ OT (x1.5)': 'OT hours (x1.5)',
  'Giờ x2': 'x2 hours',
  'TỔNG': 'TOTAL',
  'Ca làm': 'Shifts',
  'Bảng lương': 'Payroll',
  'Không tạo được file Excel': 'Could not create the Excel file',
};

final _enRules = <(RegExp, String Function(RegExpMatch))>[
  (RegExp(r'^(\d+) ca$'), (RegExpMatch m) => '${m[1]} ${m[1] == '1' ? 'shift' : 'shifts'}'),
  (RegExp(r'^(.+) · (\d+) ca$'), (RegExpMatch m) => '${m[1]} · ${m[2]} ${m[2] == '1' ? 'shift' : 'shifts'}'),
  (RegExp(r'^Vào lúc (.+) · (.+)/giờ$'), (RegExpMatch m) => 'Started ${m[1]} · ${m[2]}/hr'),
  (RegExp(r'^(.+)/giờ$'), (RegExpMatch m) => '${m[1]}/hr'),
  (RegExp(r'^Thường (\S+) · OT (\S+) · x2 (\S+)$'), (RegExpMatch m) => 'Regular ${m[1]} · OT ${m[2]} · x2 ${m[3]}'),
  (RegExp(r'^Thường (\d+h\d\d)$'), (RegExpMatch m) => 'Regular ${m[1]}'),
  (RegExp(r'^Tuần (.+)$'), (RegExpMatch m) => 'Week ${m[1]}'),
  (RegExp(r'^Hôm nay: (.+)$'), (RegExpMatch m) => 'Today: ${m[1]}'),
  (RegExp(r'^Hôm nay tính là: (.+)$'), (RegExpMatch m) => 'Today counts as: ${m[1]}'),
  (RegExp(r'^Loại ngày: (.+)$'), (RegExpMatch m) => 'Day type: ${m[1]}'),
  (RegExp(r'^Vào (\d\d:\d\d)$'), (RegExpMatch m) => 'In ${m[1]}'),
  (RegExp(r'^Ra (\d\d:\d\d)$'), (RegExpMatch m) => 'Out ${m[1]}'),
  (RegExp(r'^≈ (.+) tính đến lúc này$'), (RegExpMatch m) => '≈ ${m[1]} so far'),
  (
    RegExp(r'^Thời lượng: (\S+)( · qua nửa đêm)?$'),
    (RegExpMatch m) => 'Duration: ${m[1]}${m[2] != null ? ' · overnight' : ''}'
  ),
  (
    RegExp(r'^Còn thiếu (\S+)\. Tuần có ngày lễ T2–T6[\s\S]*$'),
    (RegExpMatch m) => '${m[1]} short. This week has a Mon–Fri holiday, so the weekend cannot make up the difference.'
  ),
  (
    RegExp(r'^Còn thiếu (\S+)\. Thứ 7, CN[\s\S]*$'),
    (RegExpMatch m) => '${m[1]} short. Sat and Sun can make up at most 8 hours per day.'
  ),
  (
    RegExp(r'^Có (\d+) bản ghi cũ[\s\S]*$'),
    (RegExpMatch m) =>
        '${m[1]} old records could not be read and are not shown. Your data has not been deleted and is still in the DB file (a copy exists as work_records_v3.db.pre_v5.bak).'
  ),
  (RegExp(r'^Xóa ca (.+)\?$'), (RegExpMatch m) => 'Delete the shift on ${m[1]}?'),
  (RegExp(r'^Bảng lương – (.+)$'), (RegExpMatch m) => 'Payroll – ${m[1]}'),
  (RegExp(r'^Trang (.+)$'), (RegExpMatch m) => 'Page ${m[1]}'),
  (
    RegExp(r'^Thu nhập chỉ là ước tính\.[\s\S]*$'),
    (RegExpMatch m) =>
        'Earnings are estimates. Each shift is recorded on the day it starts (including overnight shifts). Mon–Fri: first 8 hours x1, hours 8–12 x1.5, over 12 hours x2. Sat and Sun: the first 8 hours make up missing regular hours for the week (up to 40 hours), then x1.5, and over 12 hours x2. Holidays: x2 and never count as regular hours; if a Mon–Fri holiday falls in the week, the weekend cannot make up missing hours.'
  ),
  (
    RegExp(r'^Tuần làm việc bắt đầu thứ 2\.[\s\S]*$'),
    (RegExpMatch m) =>
        'The work week starts on Monday. Mon–Fri are weekdays, Sat and Sun are OT days (automatic). The default hourly rate only applies to new shifts.'
  ),
  (RegExp(r'^Không thể tải dữ liệu\n([\s\S]*)$'), (RegExpMatch m) => 'Could not load data\n${m[1]}'),
  (RegExp(r'^Không thể lưu ca: ([\s\S]*)$'), (RegExpMatch m) => 'Could not save the shift: ${m[1]}'),
  (RegExp(r'^Không thể lưu: ([\s\S]*)$'), (RegExpMatch m) => 'Could not save: ${m[1]}'),
  (RegExp(r'^Không thể xóa: ([\s\S]*)$'), (RegExpMatch m) => 'Could not delete: ${m[1]}'),
  (
    RegExp(r'^Không thể tạo PDF \(kiểm tra đã khai báo font trong pubspec\.yaml\): ([\s\S]*)$'),
    (RegExpMatch m) => 'Could not create the PDF (check that the font is declared in pubspec.yaml): ${m[1]}'
  ),
  (RegExp(r'^Không thể tạo file: ([\s\S]*)$'), (RegExpMatch m) => 'Could not create the file: ${m[1]}'),
  (RegExp(r'^Không thể chia sẻ file: ([\s\S]*)$'), (RegExpMatch m) => 'Could not share the file: ${m[1]}'),
];

String tr(String s) {
  if (currentLang != 'en') return s;
  final exact = _enExact[s];
  if (exact != null) return exact;
  for (final (re, build) in _enRules) {
    final match = re.firstMatch(s);
    if (match != null) return build(match);
  }
  return s;
}

/// Thay cho Text của Flutter: tự dịch chuỗi theo ngôn ngữ hiện tại.
class Text extends StatelessWidget {
  const Text(
    this.data, {
    super.key,
    this.style,
    this.textAlign,
    this.maxLines,
    this.overflow,
    this.softWrap,
  });
  final String data;
  final TextStyle? style;
  final TextAlign? textAlign;
  final int? maxLines;
  final TextOverflow? overflow;
  final bool? softWrap;

  @override
  Widget build(BuildContext context) => mat.Text(
        tr(data),
        style: style,
        textAlign: textAlign,
        maxLines: maxLines,
        overflow: overflow,
        softWrap: softWrap,
      );
}

const dayTypes = ['Thường', 'OT', 'Lễ'];
const kDefaultRate = 25.0;

const _tealMid = Color(0xFF0F766E);
const _tealDark = Color(0xFF0B4F4A);
const _surfaceBg = Color(0xFFF3F6F8);
const _muted = Color(0xFF64748B);
String _fmtStart(DateTime d) => '${_two(d.hour)}:${_two(d.minute)} ${_two(d.day)}/${_two(d.month)}';

final _moneyFmt = NumberFormat.currency(locale: 'en_US', symbol: r'$', decimalDigits: 2);
final _rateFmt = NumberFormat('0.##', 'en_US');
final _dateFmt = DateFormat('yyyy-MM-dd');
final _dmyFmt = DateFormat('dd/MM/yyyy');
final _timeFmt = DateFormat('HH:mm');

String money(double v) => _moneyFmt.format(v);

String fmtHours(double hours) {
  final m = (hours * 60).round();
  return '${m ~/ 60}h${(m % 60).toString().padLeft(2, '0')}';
}

String _two(int n) => n.toString().padLeft(2, '0');

DateTime floorMinute(DateTime d) => DateTime(d.year, d.month, d.day, d.hour, d.minute);
DateTime dateOnly(DateTime d) => DateTime(d.year, d.month, d.day);

DateTime startOfWeek(DateTime d, int weekStart) {
  final day = dateOnly(d);
  return DateTime(day.year, day.month, day.day - (day.weekday - weekStart + 7) % 7);
}

/// Loại ngày tự suy ra: ngày lễ (đánh dấu trong cài đặt) > thứ 7/CN (OT) > ngày thường.
String dayTypeOf(DateTime d, Set<String> holidays) {
  if (holidays.contains(_dateFmt.format(d))) return 'Lễ';
  return d.weekday >= DateTime.saturday ? 'OT' : 'Thường';
}

/// Nhận cả "23.5", "23,5", "1,234.50", "1.234,50".
double? parseNumber(String input) {
  var s = input.trim().replaceAll(RegExp(r'[\s$]'), '');
  if (s.isEmpty) return null;
  final hasComma = s.contains(',');
  final hasDot = s.contains('.');
  if (hasComma && hasDot) {
    final commaIsDecimal = s.lastIndexOf(',') > s.lastIndexOf('.');
    s = commaIsDecimal ? s.replaceAll('.', '').replaceAll(',', '.') : s.replaceAll(',', '');
  } else if (hasComma) {
    s = s.replaceAll(',', '.');
  }
  final v = double.tryParse(s);
  return (v == null || !v.isFinite || v < 0) ? null : v;
}

Color typeColor(String t) => switch (t) {
      'Lễ' => Colors.deepOrange,
      'OT' => Colors.indigo,
      _ => Colors.teal,
    };

IconData typeIcon(String t) => switch (t) {
      'Lễ' => Icons.celebration,
      'OT' => Icons.more_time,
      _ => Icons.work_outline,
    };

bool _near(double a, double b) => (a - b).abs() < 1e-6;

// ════════════════════════════════════════════════════════════════════════════
//  Model
// ════════════════════════════════════════════════════════════════════════════

class WorkRecord {
  const WorkRecord({
    this.id,
    required this.start,
    required this.end,
    required this.dayType,
    required this.hourlyRate,
    this.totalHours = 0,
    this.regularHours = 0,
    this.otHours = 0,
    this.doubleHours = 0,
    this.earned = 0,
  });

  final int? id;
  final DateTime start;
  final DateTime end;
  final String dayType;
  final double hourlyRate;
  final double totalHours;
  final double regularHours;
  final double otHours; // x1.5
  final double doubleHours; // x2
  final double earned;

  int get minutes => end.difference(start).inMinutes;
  bool get overnight => dateOnly(end) != dateOnly(start);

  WorkRecord withId(int newId) => WorkRecord(
        id: newId,
        start: start,
        end: end,
        dayType: dayType,
        hourlyRate: hourlyRate,
        totalHours: totalHours,
        regularHours: regularHours,
        otHours: otHours,
        doubleHours: doubleHours,
        earned: earned,
      );

  WorkRecord withPay({
    required double total,
    required double regular,
    required double ot,
    required double dbl,
    required double earned,
    required String dayType,
  }) =>
      WorkRecord(
        id: id,
        start: start,
        end: end,
        dayType: dayType,
        hourlyRate: hourlyRate,
        totalHours: total,
        regularHours: regular,
        otHours: ot,
        doubleHours: dbl,
        earned: earned,
      );

  bool samePay(WorkRecord o) =>
      dayType == o.dayType &&
      _near(totalHours, o.totalHours) &&
      _near(regularHours, o.regularHours) &&
      _near(otHours, o.otHours) &&
      _near(doubleHours, o.doubleHours) &&
      (earned - o.earned).abs() < 0.005;

  // Vẫn ghi dateStr/checkInStr/checkOutStr để tương thích cột cũ (NOT NULL).
  Map<String, Object?> toMap() => {
        'dateStr': _dateFmt.format(start),
        'checkInStr': _timeFmt.format(start),
        'checkOutStr': _timeFmt.format(end),
        'startMs': start.millisecondsSinceEpoch,
        'endMs': end.millisecondsSinceEpoch,
        'totalHours': totalHours,
        'regularHours': regularHours,
        'otHours': otHours,
        'doubleTimeHours': doubleHours,
        'dayType': dayType,
        'hourlyRate': hourlyRate,
        'earned': earned,
      };

  static WorkRecord? fromMap(Map<String, Object?> m) {
    double n(String k) => (m[k] as num?)?.toDouble() ?? 0;
    final startMs = (m['startMs'] as num?)?.toInt() ?? 0;
    final endMs = (m['endMs'] as num?)?.toInt() ?? 0;
    if (startMs <= 0 || endMs <= startMs) return null;
    final type = m['dayType'] as String? ?? 'Thường';
    return WorkRecord(
      id: m['id'] as int?,
      start: DateTime.fromMillisecondsSinceEpoch(startMs),
      end: DateTime.fromMillisecondsSinceEpoch(endMs),
      dayType: dayTypes.contains(type) ? type : 'Thường',
      hourlyRate: n('hourlyRate'),
      totalHours: n('totalHours'),
      regularHours: n('regularHours'),
      otHours: n('otHours'),
      doubleHours: n('doubleTimeHours'),
      earned: n('earned'),
    );
  }
}

class RecordDraft {
  const RecordDraft({
    this.id,
    required this.start,
    required this.end,
    required this.rate,
  });
  final int? id;
  final DateTime start;
  final DateTime end;
  final double rate;
}

class PeriodSummary {
  const PeriodSummary({
    this.earned = 0,
    this.total = 0,
    this.regular = 0,
    this.ot = 0,
    this.dbl = 0,
    this.shifts = 0,
  });
  final double earned, total, regular, ot, dbl;
  final int shifts;
}

PeriodSummary summarize(Iterable<WorkRecord> records, DateTime from, DateTime to) {
  double e = 0, t = 0, r = 0, o = 0, d = 0;
  var n = 0;
  for (final x in records) {
    if (x.start.isBefore(from) || !x.start.isBefore(to)) continue;
    e += x.earned;
    t += x.totalHours;
    r += x.regularHours;
    o += x.otHours;
    d += x.doubleHours;
    n++;
  }
  return PeriodSummary(earned: e, total: t, regular: r, ot: o, dbl: d, shifts: n);
}

// ════════════════════════════════════════════════════════════════════════════
//  Tính lương (hàm thuần, dễ unit test)
//
//  Mỗi ca được ghi vào NGÀY BẮT ĐẦU. Tuần bắt đầu thứ 2.
//  Loại ngày tự suy ra: Lễ (đánh dấu trong cài đặt) > thứ 7/CN = OT > Thường.
//
//   • Ngày thường (T2–T6): 8 giờ đầu x1 (tối đa 40 giờ thường/tuần, vượt → x1.5);
//     giờ 8–12 x1.5; trên 12 giờ x2. Giờ OT trong ngày KHÔNG dùng để bù giờ thường.
//   • Thứ 7/CN (OT): x1.5, giờ 8–12 x1.5, trên 12 giờ/ngày x2. Nếu T2–T6 chưa đủ
//     40 giờ thường thì tối đa 8 giờ đầu của mỗi ngày cuối tuần được bù vào x1.
//   • Lễ: cả ca x2, không tính vào giờ thường. Tuần có ngày lễ rơi vào T2–T6
//     thì cuối tuần KHÔNG được bù giờ thường (vẫn tính OT bình thường).
// ════════════════════════════════════════════════════════════════════════════

abstract final class PayCalculator {
  static const _dailyRegular = 8 * 60;
  static const _dailyOt = 12 * 60;
  static const _weeklyRegular = 40 * 60;

  static List<WorkRecord> calculate(List<WorkRecord> input, Set<String> holidays) {
    final sorted = [...input]..sort((a, b) {
        final c = a.start.compareTo(b.start);
        return c != 0 ? c : (a.id ?? 0).compareTo(b.id ?? 0);
      });
    final daily = <DateTime, int>{};
    final weekly = <DateTime, int>{};
    return [for (final r in sorted) _one(r, holidays, daily, weekly)];
  }

  /// Tuần (bắt đầu thứ 2) có ngày lễ rơi vào T2–T6?
  static bool _weekdayHolidayIn(DateTime monday, Set<String> holidays) {
    for (var i = 0; i < 5; i++) {
      final d = DateTime(monday.year, monday.month, monday.day + i);
      if (holidays.contains(_dateFmt.format(d))) return true;
    }
    return false;
  }

  static WorkRecord _one(
    WorkRecord r,
    Set<String> holidays,
    Map<DateTime, int> daily,
    Map<DateTime, int> weekly,
  ) {
    final day = dateOnly(r.start);
    final week = startOfWeek(r.start, DateTime.monday);
    final type = dayTypeOf(r.start, holidays);
    var remaining = r.minutes;
    var reg = 0, ot = 0, dbl = 0;

    if (type == 'Lễ') {
      // Cả ca x2, độc lập: không cộng vào bộ đếm ngày/tuần.
      dbl = remaining;
      remaining = 0;
    } else {
      final weekend = type == 'OT';
      final canFill = weekend && !_weekdayHolidayIn(week, holidays);

      while (remaining > 0) {
        final used = daily[day] ?? 0;
        final int room;
        if (used >= _dailyOt) {
          room = remaining;
        } else if (used < _dailyRegular) {
          room = _dailyRegular - used;
        } else {
          room = _dailyOt - used;
        }
        final chunk = math.min(remaining, room);

        if (used >= _dailyOt) {
          dbl += chunk;
        } else if (used >= _dailyRegular) {
          ot += chunk; // giờ 8–12 trong ngày: x1.5
        } else if (weekend) {
          // 8 giờ đầu của ngày cuối tuần: bù vào giờ thường còn thiếu của tuần
          // (tối đa 40 giờ) trước, phần còn lại x1.5.
          var fill = 0;
          if (canFill) {
            final usedWeek = weekly[week] ?? 0;
            fill = math.min(chunk, math.max(0, _weeklyRegular - usedWeek));
            reg += fill;
            weekly[week] = usedWeek + fill;
          }
          ot += chunk - fill;
        } else {
          // Ngày thường, 8 giờ đầu: giờ thường, tối đa 40 giờ/tuần, vượt → OT.
          final usedWeek = weekly[week] ?? 0;
          final regular = math.min(chunk, math.max(0, _weeklyRegular - usedWeek));
          reg += regular;
          ot += chunk - regular;
          weekly[week] = usedWeek + regular;
        }
        daily[day] = used + chunk;
        remaining -= chunk;
      }
    }

    final raw = (reg + ot * 1.5 + dbl * 2) / 60.0 * r.hourlyRate;
    return r.withPay(
      total: r.minutes / 60.0,
      regular: reg / 60.0,
      ot: ot / 60.0,
      dbl: dbl / 60.0,
      earned: (raw * 100).roundToDouble() / 100,
      dayType: type,
    );
  }
}

/// Ước tính tiền của ca đang làm tính đến `now` (chỉ tính lại trong tuần của ca).
double estimateActive(
  List<WorkRecord> records,
  ActiveShift a,
  Set<String> holidays,
  DateTime now,
) {
  final end = floorMinute(now);
  if (!end.isAfter(a.start)) return 0;
  final week = startOfWeek(a.start, DateTime.monday);
  final next = DateTime(week.year, week.month, week.day + 7);
  const probeId = -1;
  final raw = [
    for (final r in records)
      if (!r.start.isBefore(week) && r.start.isBefore(next)) r,
    WorkRecord(id: probeId, start: a.start, end: end, dayType: 'Thường', hourlyRate: a.rate),
  ];
  return PayCalculator.calculate(raw, holidays).firstWhere((r) => r.id == probeId).earned;
}

// ════════════════════════════════════════════════════════════════════════════
//  Lưu trữ
// ════════════════════════════════════════════════════════════════════════════

class AppDatabase {
  AppDatabase._();
  static final instance = AppDatabase._();

  Future<Database>? _open;

  Future<Database> get db async {
    try {
      return await (_open ??= _openDb());
    } catch (_) {
      _open = null;
      rethrow;
    }
  }

  Future<Database> _openDb() async {
    final dir = await getDatabasesPath();
    final path = p.join(dir, 'work_records_v3.db');
    await _backupOnce(path);
    final db = await openDatabase(
      path,
      version: 5,
      onCreate: _onCreate,
      onUpgrade: _onUpgrade,
    );
    await _importLegacy(db, dir);
    return db;
  }

  /// Nhập một lần dữ liệu từ các file DB tên cũ (nếu có). Không bao giờ làm
  /// app lỗi khi mở: thất bại thì bỏ qua và thử lại ở lần mở sau.
  Future<void> _importLegacy(Database db, String dir) async {
    final prefs = await SharedPreferences.getInstance();
    final rate = prefs.getDouble('hourlyRate') ?? kDefaultRate;

    for (final name in const ['work_records.db', 'work_records_v2.db']) {
      final flag = 'legacyImported:$name';
      if (prefs.getBool(flag) == true) continue;
      try {
        final legacyPath = p.join(dir, name);
        if (!await databaseExists(legacyPath)) continue;

        final old = await openDatabase(legacyPath, readOnly: true);
        try {
          final rows = await old.query('records');
          final batch = db.batch();
          for (final m in rows) {
            final start = DateTime.tryParse('${m['dateStr']} ${m['checkInStr']}:00');
            final end0 = DateTime.tryParse('${m['dateStr']} ${m['checkOutStr']}:00');
            if (start == null || end0 == null) continue;
            var end = end0;
            if (!end.isAfter(start)) {
              end = DateTime(end.year, end.month, end.day + 1, end.hour, end.minute);
            }
            final type = m['dayType'] as String?;
            final oldRate = (m['hourlyRate'] as num?)?.toDouble();
            batch.insert(
              'records',
              WorkRecord(
                start: start,
                end: end,
                dayType: dayTypes.contains(type) ? type! : 'Thường',
                hourlyRate: (oldRate != null && oldRate > 0) ? oldRate : rate,
              ).toMap(),
            );
          }
          await batch.commit(noResult: true);
        } finally {
          await old.close();
        }
        await prefs.setBool(flag, true);
      } catch (_) {
        // bỏ qua, thử lại lần mở sau
      }
    }
  }

  /// Chép file DB hiện có sang bản .bak một lần duy nhất, trước khi nâng cấp
  /// cấu trúc. Nếu có sự cố, dữ liệu gốc vẫn còn nguyên.
  Future<void> _backupOnce(String path) async {
    try {
      final src = File(path);
      final bak = File('$path.pre_v5.bak');
      if (await src.exists() && !await bak.exists()) await src.copy(bak.path);
    } catch (_) {}
  }

  Future<void> _onCreate(Database db, int version) async {
    await db.execute('''
      CREATE TABLE records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dateStr TEXT NOT NULL,
        checkInStr TEXT NOT NULL,
        checkOutStr TEXT NOT NULL,
        startMs INTEGER NOT NULL DEFAULT 0,
        endMs INTEGER NOT NULL DEFAULT 0,
        totalHours REAL NOT NULL DEFAULT 0,
        regularHours REAL NOT NULL DEFAULT 0,
        otHours REAL NOT NULL DEFAULT 0,
        doubleTimeHours REAL NOT NULL DEFAULT 0,
        dayType TEXT NOT NULL DEFAULT 'Thường',
        hourlyRate REAL NOT NULL DEFAULT 0,
        earned REAL NOT NULL DEFAULT 0
      )
    ''');
    await db.execute('CREATE INDEX idx_records_start ON records(startMs)');
  }

  Future<void> _onUpgrade(Database db, int oldVersion, int newVersion) async {
    final cols = (await db.rawQuery('PRAGMA table_info(records)'))
        .map((c) => c['name'] as String)
        .toSet();
    Future<void> add(String name, String type) async {
      if (!cols.contains(name)) {
        await db.execute('ALTER TABLE records ADD COLUMN $name $type NOT NULL DEFAULT 0');
      }
    }

    await add('doubleTimeHours', 'REAL');
    await add('hourlyRate', 'REAL');
    await add('startMs', 'INTEGER');
    await add('endMs', 'INTEGER');

    // Chuyển bản ghi cũ (chuỗi ngày + HH:mm) sang mốc thời gian đầy đủ.
    final rows = await db.query(
      'records',
      columns: ['id', 'dateStr', 'checkInStr', 'checkOutStr'],
      where: 'startMs = 0',
    );
    final batch = db.batch();
    for (final r in rows) {
      final start = DateTime.tryParse('${r['dateStr']} ${r['checkInStr']}:00');
      final end0 = DateTime.tryParse('${r['dateStr']} ${r['checkOutStr']}:00');
      if (start == null || end0 == null) continue;
      var end = end0;
      if (!end.isAfter(start)) {
        end = DateTime(end.year, end.month, end.day + 1, end.hour, end.minute);
      }
      batch.update(
        'records',
        {'startMs': start.millisecondsSinceEpoch, 'endMs': end.millisecondsSinceEpoch},
        where: 'id = ?',
        whereArgs: [r['id']],
      );
    }
    await batch.commit(noResult: true);
    await db.execute('CREATE INDEX IF NOT EXISTS idx_records_start ON records(startMs)');
  }

  Future<List<WorkRecord>> getAll() async {
    final rows = await (await db).query('records', orderBy: 'startMs DESC, id DESC');
    return rows.map(WorkRecord.fromMap).whereType<WorkRecord>().toList();
  }

  /// Số dòng còn trong DB nhưng không đọc được (giờ/ngày hỏng). Không bị xóa.
  Future<int> countUnreadable() async {
    final rows = await (await db).rawQuery(
      'SELECT COUNT(*) FROM records WHERE startMs <= 0 OR endMs <= startMs',
    );
    return Sqflite.firstIntValue(rows) ?? 0;
  }

  Future<int> insert(WorkRecord r) async => (await db).insert('records', r.toMap());

  Future<void> update(WorkRecord r) async {
    if (r.id == null) return;
    await (await db).update('records', r.toMap(), where: 'id = ?', whereArgs: [r.id]);
  }

  Future<void> delete(int id) async =>
      (await db).delete('records', where: 'id = ?', whereArgs: [id]);

  Future<void> updateAll(List<WorkRecord> records) async {
    if (records.isEmpty) return;
    final batch = (await db).batch();
    for (final r in records) {
      if (r.id == null) continue;
      batch.update('records', r.toMap(), where: 'id = ?', whereArgs: [r.id]);
    }
    await batch.commit(noResult: true);
  }
}

class AppSettings {
  const AppSettings({this.hourlyRate = kDefaultRate, this.holidays = const {}, this.lang = 'vi'});
  final double hourlyRate;
  final Set<String> holidays; // 'yyyy-MM-dd'
  final String lang; // 'vi' hoặc 'en'

  /// Quy tắc cố định: tuần làm việc bắt đầu từ thứ 2.
  int get weekStart => DateTime.monday;

  static Future<AppSettings> load() async {
    final prefs = await SharedPreferences.getInstance();
    return AppSettings(
      hourlyRate: prefs.getDouble('hourlyRate') ?? kDefaultRate,
      holidays: (prefs.getStringList('holidays') ?? const <String>[]).toSet(),
      lang: supportedLangs.containsKey(prefs.getString('lang')) ? prefs.getString('lang')! : 'vi',
    );
  }

  Future<void> save() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setDouble('hourlyRate', hourlyRate);
    await prefs.setStringList('holidays', holidays.toList()..sort());
    await prefs.setString('lang', lang);
  }
}

class ActiveShift {
  const ActiveShift({required this.start, required this.dayType, required this.rate});
  final DateTime start;
  final String dayType;
  final double rate;

  static Future<ActiveShift?> load(double fallbackRate) async {
    final prefs = await SharedPreferences.getInstance();
    final ms = prefs.getInt('activeCheckInMillis');
    if (ms == null) return null;
    final type = prefs.getString('activeDayType') ?? 'Thường';
    return ActiveShift(
      start: DateTime.fromMillisecondsSinceEpoch(ms),
      dayType: dayTypes.contains(type) ? type : 'Thường',
      rate: prefs.getDouble('activeRate') ?? fallbackRate,
    );
  }

  Future<void> save() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setInt('activeCheckInMillis', start.millisecondsSinceEpoch);
    await prefs.setString('activeDayType', dayType);
    await prefs.setDouble('activeRate', rate);
  }

  static Future<void> clear() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.remove('activeCheckInMillis');
    await prefs.remove('activeDayType');
    await prefs.remove('activeRate');
  }
}

String buildCsv(List<WorkRecord> newestFirst) {
  String cell(Object? v) => '"${'$v'.replaceAll('"', '""')}"';
  final rows = <List<Object?>>[
    [
      tr('Ngày'), tr('Giờ vào'), tr('Giờ ra'), tr('Loại ngày'), tr('Tổng giờ'), tr('Giờ thường'),
      tr('Giờ OT (x1.5)'), tr('Giờ x2'), tr('Lương/giờ (USD)'), tr('Tiền (USD)')
    ],
    for (final r in newestFirst.reversed)
      [
        _dateFmt.format(r.start),
        _timeFmt.format(r.start),
        _timeFmt.format(r.end) + (r.overnight ? ' (+1)' : ''),
        dayLabel(r.dayType),
        r.totalHours.toStringAsFixed(2),
        r.regularHours.toStringAsFixed(2),
        r.otHours.toStringAsFixed(2),
        r.doubleHours.toStringAsFixed(2),
        r.hourlyRate.toStringAsFixed(2),
        r.earned.toStringAsFixed(2),
      ],
  ];
  return '\uFEFF${rows.map((row) => row.map(cell).join(',')).join('\r\n')}';
}

// ════════════════════════════════════════════════════════════════════════════
//  Controller (toàn bộ logic nghiệp vụ, UI chỉ lắng nghe)
// ════════════════════════════════════════════════════════════════════════════

class AppController extends ChangeNotifier {
  AppController(this._db);
  final AppDatabase _db;

  List<WorkRecord> records = const []; // mới nhất trước
  AppSettings settings = const AppSettings();
  ActiveShift? active;
  bool loading = true;
  String? loadError;
  int skipped = 0; // bản ghi cũ không đọc được (vẫn còn trong file DB)

  Future<void> load() async {
    loading = true;
    loadError = null;
    notifyListeners();
    try {
      settings = await AppSettings.load();
      active = await ActiveShift.load(settings.hourlyRate);
      skipped = await _db.countUnreadable();
      await _recalculate(await _db.getAll());
    } catch (e) {
      loadError = '$e';
    }
    loading = false;
    notifyListeners();
  }

  Future<void> refresh() async {
    try {
      await _recalculate(await _db.getAll());
    } catch (e) {
      loadError = '$e';
      notifyListeners();
    }
  }

  /// Tính lại toàn bộ và chỉ ghi những bản ghi có thay đổi.
  Future<void> _recalculate(List<WorkRecord> raw) async {
    final calculated = PayCalculator.calculate(raw, settings.holidays);
    final before = {for (final r in raw) r.id: r};
    final changed = [
      for (final r in calculated)
        if (before[r.id]?.samePay(r) != true) r
    ];
    await _db.updateAll(changed);
    records = calculated.reversed.toList();
    notifyListeners();
  }

  String? validate(RecordDraft d) {
    if (!d.end.isAfter(d.start)) return 'Giờ ra phải sau giờ vào.';
    if (d.end.difference(d.start) > const Duration(hours: 24)) {
      return 'Một ca tối đa 24 giờ.';
    }
    final clash = records.any(
      (r) => r.id != d.id && r.start.isBefore(d.end) && r.end.isAfter(d.start),
    );
    return clash ? 'Ca này trùng giờ với một ca đã có.' : null;
  }

  Future<String?> saveDraft(RecordDraft d) async {
    final err = validate(d);
    if (err != null) return err;
    try {
      final base = WorkRecord(
        id: d.id,
        start: d.start,
        end: d.end,
        dayType: 'Thường', // loại ngày thật do PayCalculator suy ra
        hourlyRate: d.rate,
      );
      final int id;
      if (d.id == null) {
        id = await _db.insert(base);
      } else {
        await _db.update(base);
        id = d.id!;
      }
      await _recalculate([...records.where((r) => r.id != id), base.withId(id)]);
      return null;
    } catch (e) {
      return 'Không thể lưu: $e';
    }
  }

  Future<String?> delete(WorkRecord r) async {
    try {
      await _db.delete(r.id!);
      await _recalculate(records.where((x) => x.id != r.id).toList());
      return null;
    } catch (e) {
      return 'Không thể xóa: $e';
    }
  }

  Future<void> updateSettings(AppSettings s) async {
    settings = s;
    await s.save();
    await _recalculate(List.of(records)); // đổi ngày lễ → tính lại
  }

  String dayTypeFor(DateTime d) => dayTypeOf(d, settings.holidays);

  Future<void> checkIn(double rate) async {
    final start = floorMinute(DateTime.now());
    final shift = ActiveShift(start: start, dayType: dayTypeFor(start), rate: rate);
    await shift.save();
    active = shift;
    notifyListeners();
  }

  Future<String?> checkOut() async {
    final shift = active;
    if (shift == null) return null;
    final end = floorMinute(DateTime.now());
    if (!end.isAfter(shift.start)) {
      return 'Ca làm cần ít nhất 1 phút. Bấm "Hủy ca" nếu bạn bấm nhầm.';
    }
    try {
      final base = WorkRecord(
        start: shift.start,
        end: end,
        dayType: shift.dayType,
        hourlyRate: shift.rate,
      );
      final id = await _db.insert(base);
      await ActiveShift.clear();
      active = null;
      await _recalculate([...records, base.withId(id)]);
      return null;
    } catch (e) {
      return 'Không thể lưu ca: $e';
    }
  }

  Future<void> discardActive() async {
    await ActiveShift.clear();
    active = null;
    notifyListeners();
  }
}

// ════════════════════════════════════════════════════════════════════════════
//  Xuất file: CSV, Excel (.xlsx), PDF
// ════════════════════════════════════════════════════════════════════════════

enum ExportRange { week, month, lastMonth, year, all }

enum ExportFormat { csv, xlsx, pdf }

const _rangeLabels = {
  ExportRange.week: 'Tuần này',
  ExportRange.month: 'Tháng này',
  ExportRange.lastMonth: 'Tháng trước',
  ExportRange.year: 'Năm nay',
  ExportRange.all: 'Tất cả',
};

/// (từ, đến) với đến là mốc không bao gồm; null nghĩa là tất cả.
(DateTime, DateTime)? exportBounds(ExportRange r, DateTime now) {
  switch (r) {
    case ExportRange.week:
      final w = startOfWeek(now, DateTime.monday);
      return (w, DateTime(w.year, w.month, w.day + 7));
    case ExportRange.month:
      return (DateTime(now.year, now.month), DateTime(now.year, now.month + 1));
    case ExportRange.lastMonth:
      return (DateTime(now.year, now.month - 1), DateTime(now.year, now.month));
    case ExportRange.year:
      return (DateTime(now.year), DateTime(now.year + 1));
    case ExportRange.all:
      return null;
  }
}

List<WorkRecord> recordsInRange(List<WorkRecord> all, ExportRange r) {
  final b = exportBounds(r, DateTime.now());
  if (b == null) return all;
  return [
    for (final x in all)
      if (!x.start.isBefore(b.$1) && x.start.isBefore(b.$2)) x
  ];
}

/// `newestFirst` là danh sách đã lọc, mới nhất trước.
String rangeSubtitle(ExportRange r, List<WorkRecord> newestFirst) {
  final b = exportBounds(r, DateTime.now());
  if (b != null) {
    final last = DateTime(b.$2.year, b.$2.month, b.$2.day - 1);
    return '${_dmyFmt.format(b.$1)} – ${_dmyFmt.format(last)}';
  }
  if (newestFirst.isEmpty) return '';
  return '${_dmyFmt.format(newestFirst.last.start)} – ${_dmyFmt.format(newestFirst.first.start)}';
}

/// Tổng hợp theo tuần; `asc` phải sắp xếp từ cũ đến mới.
List<(DateTime, PeriodSummary)> weeklySummaries(List<WorkRecord> asc) {
  final out = <(DateTime, PeriodSummary)>[];
  DateTime? current;
  for (final r in asc) {
    final w = startOfWeek(r.start, DateTime.monday);
    if (current != w) {
      current = w;
      out.add((w, summarize(asc, w, DateTime(w.year, w.month, w.day + 7))));
    }
  }
  return out;
}

String _weekLabel(DateTime monday) {
  final sun = DateTime(monday.year, monday.month, monday.day + 6);
  return '${_two(monday.day)}/${_two(monday.month)} – ${_two(sun.day)}/${_two(sun.month)}/${sun.year}';
}

// ── Excel ───────────────────────────────────────────────────────────────────

List<int> buildXlsx(List<WorkRecord> newestFirst) {
  final asc = newestFirst.reversed.toList();
  final book = xl.Excel.createExcel();
  book.rename('Sheet1', tr('Ca làm'));
  final shifts = book[tr('Ca làm')];
  final weeks = book[tr('Theo tuần')];

  final head = xl.CellStyle(
    bold: true,
    fontColorHex: xl.ExcelColor.white,
    backgroundColorHex: xl.ExcelColor.fromHexString('#0F766E'),
    horizontalAlign: xl.HorizontalAlign.Center,
  );
  final bold = xl.CellStyle(bold: true);

  void put(xl.Sheet sh, int col, int row, xl.CellValue v, [xl.CellStyle? st]) => sh.updateCell(
        xl.CellIndex.indexByColumnRow(columnIndex: col, rowIndex: row),
        v,
        cellStyle: st,
      );
  xl.CellValue t(String v) => xl.TextCellValue(v);
  xl.CellValue n(double v) => xl.DoubleCellValue(double.parse(v.toStringAsFixed(2)));

  // Sheet 1: chi tiết từng ca
  final h1 = [
    tr('Ngày'), tr('Thứ'), tr('Giờ vào'), tr('Giờ ra'), tr('Loại ngày'), tr('Tổng giờ'),
    tr('Giờ thường'), tr('Giờ OT (x1.5)'), tr('Giờ x2'), tr('Lương/giờ (USD)'), tr('Tiền (USD)')
  ];
  for (var c = 0; c < h1.length; c++) {
    put(shifts, c, 0, t(h1[c]), head);
    shifts.setColumnWidth(c, c == 0 ? 14 : 13);
  }
  var row = 1;
  double sTotal = 0, sReg = 0, sOt = 0, sDbl = 0, sEarned = 0;
  for (final r in asc) {
    put(shifts, 0, row, t(_dmyFmt.format(r.start)));
    put(shifts, 1, row, t(_weekdayShort[r.start.weekday - 1]));
    put(shifts, 2, row, t(_timeFmt.format(r.start)));
    put(shifts, 3, row, t('${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}'));
    put(shifts, 4, row, t(dayLabel(r.dayType)));
    put(shifts, 5, row, n(r.totalHours));
    put(shifts, 6, row, n(r.regularHours));
    put(shifts, 7, row, n(r.otHours));
    put(shifts, 8, row, n(r.doubleHours));
    put(shifts, 9, row, n(r.hourlyRate));
    put(shifts, 10, row, n(r.earned));
    sTotal += r.totalHours;
    sReg += r.regularHours;
    sOt += r.otHours;
    sDbl += r.doubleHours;
    sEarned += r.earned;
    row++;
  }
  put(shifts, 0, row, t(tr('TỔNG')), bold);
  put(shifts, 5, row, n(sTotal), bold);
  put(shifts, 6, row, n(sReg), bold);
  put(shifts, 7, row, n(sOt), bold);
  put(shifts, 8, row, n(sDbl), bold);
  put(shifts, 10, row, n(sEarned), bold);

  // Sheet 2: tổng hợp theo tuần
  final h2 = [tr('Tuần'), tr('Số ca'), tr('Tổng giờ'), tr('Giờ thường'), tr('Giờ OT (x1.5)'), tr('Giờ x2'), tr('Tiền (USD)')];
  for (var c = 0; c < h2.length; c++) {
    put(weeks, c, 0, t(h2[c]), head);
    weeks.setColumnWidth(c, c == 0 ? 24 : 14);
  }
  var wr = 1;
  for (final (monday, sum) in weeklySummaries(asc)) {
    put(weeks, 0, wr, t(_weekLabel(monday)));
    put(weeks, 1, wr, xl.IntCellValue(sum.shifts));
    put(weeks, 2, wr, n(sum.total));
    put(weeks, 3, wr, n(sum.regular));
    put(weeks, 4, wr, n(sum.ot));
    put(weeks, 5, wr, n(sum.dbl));
    put(weeks, 6, wr, n(sum.earned));
    wr++;
  }

  final bytes = book.save();
  if (bytes == null) throw StateError(tr('Không tạo được file Excel'));
  return bytes;
}

// ── PDF ─────────────────────────────────────────────────────────────────────

/// Chạy hoàn toàn offline: dùng font Liberation Sans đóng gói sẵn trong assets (có tiếng Việt).
Future<List<int>> buildPdf(List<WorkRecord> newestFirst, String title, String subtitle) async {
  final asc = newestFirst.reversed.toList();
  final regular = pw.Font.ttf(await rootBundle.load('assets/fonts/LiberationSans-Regular.ttf'));
  final boldFont = pw.Font.ttf(await rootBundle.load('assets/fonts/LiberationSans-Bold.ttf'));
  final doc = pw.Document(theme: pw.ThemeData.withFont(base: regular, bold: boldFont));

  final teal = pdf.PdfColor.fromHex('#0F766E');
  final grey = pdf.PdfColor.fromHex('#64748B');
  final light = pdf.PdfColor.fromHex('#F1F5F9');
  final sum = summarize(asc, DateTime(1970), DateTime(3000));

  pw.Widget stat(String label, String value) => pw.Expanded(
        child: pw.Container(
          padding: const pw.EdgeInsets.all(10),
          decoration: pw.BoxDecoration(color: light),
          child: pw.Column(
            crossAxisAlignment: pw.CrossAxisAlignment.start,
            children: [
              pw.Text(label, style: pw.TextStyle(fontSize: 9, color: grey)),
              pw.SizedBox(height: 3),
              pw.Text(value, style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
            ],
          ),
        ),
      );

  pw.Widget table(List<String> headers, List<List<String>> data) => pw.TableHelper.fromTextArray(
        headers: headers,
        data: data,
        headerStyle: pw.TextStyle(fontSize: 9, fontWeight: pw.FontWeight.bold, color: pdf.PdfColors.white),
        headerDecoration: pw.BoxDecoration(color: teal),
        cellStyle: const pw.TextStyle(fontSize: 9),
        oddRowDecoration: pw.BoxDecoration(color: light),
        cellPadding: const pw.EdgeInsets.symmetric(horizontal: 5, vertical: 4),
        cellAlignments: {
          0: pw.Alignment.centerLeft,
          for (var i = 1; i < headers.length; i++) i: pw.Alignment.centerRight,
        },
      );

  doc.addPage(
    pw.MultiPage(
      pageFormat: pdf.PdfPageFormat.a4,
      margin: const pw.EdgeInsets.all(32),
      footer: (ctx) => pw.Align(
        alignment: pw.Alignment.centerRight,
        child: pw.Text(
          tr('Trang ${ctx.pageNumber}/${ctx.pagesCount}'),
          style: pw.TextStyle(fontSize: 9, color: grey),
        ),
      ),
      build: (ctx) => [
        pw.Text(title, style: pw.TextStyle(fontSize: 20, fontWeight: pw.FontWeight.bold, color: teal)),
        if (subtitle.isNotEmpty) pw.Text(subtitle, style: pw.TextStyle(fontSize: 10, color: grey)),
        pw.SizedBox(height: 14),
        pw.Row(
          children: [
            stat(tr('Tổng tiền'), money(sum.earned)),
            pw.SizedBox(width: 8),
            stat(tr('Tổng giờ'), fmtHours(sum.total)),
            pw.SizedBox(width: 8),
            stat(tr('Giờ thường'), fmtHours(sum.regular)),
            pw.SizedBox(width: 8),
            stat('OT x1.5', fmtHours(sum.ot)),
            pw.SizedBox(width: 8),
            stat('x2', fmtHours(sum.dbl)),
          ],
        ),
        pw.SizedBox(height: 18),
        pw.Text(tr('Theo tuần'), style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
        pw.SizedBox(height: 6),
        table(
          [tr('Tuần'), tr('Số ca'), tr('Tổng giờ'), tr('Thường'), 'OT x1.5', 'x2', tr('Tiền')],
          [
            for (final (monday, w) in weeklySummaries(asc))
              [
                _weekLabel(monday),
                '${w.shifts}',
                fmtHours(w.total),
                fmtHours(w.regular),
                fmtHours(w.ot),
                fmtHours(w.dbl),
                money(w.earned),
              ],
          ],
        ),
        pw.SizedBox(height: 18),
        pw.Text(tr('Chi tiết ca làm'), style: pw.TextStyle(fontSize: 12, fontWeight: pw.FontWeight.bold)),
        pw.SizedBox(height: 6),
        table(
          [tr('Ngày'), tr('Vào'), tr('Ra'), tr('Loại'), tr('Thường'), 'OT x1.5', 'x2', tr('Lương/giờ'), tr('Tiền')],
          [
            for (final r in asc)
              [
                '${_weekdayShort[r.start.weekday - 1]} ${_dmyFmt.format(r.start)}',
                _timeFmt.format(r.start),
                '${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}',
                dayLabel(r.dayType),
                fmtHours(r.regularHours),
                fmtHours(r.otHours),
                fmtHours(r.doubleHours),
                money(r.hourlyRate),
                money(r.earned),
              ],
          ],
        ),
      ],
    ),
  );
  return doc.save();
}

// ── Tạo file + chia sẻ ──────────────────────────────────────────────────────

Future<({List<int> bytes, String name, String mime})> buildExport(
  List<WorkRecord> newestFirst,
  ExportRange range,
  ExportFormat format,
) async {
  final stamp = _dateFmt.format(DateTime.now());
  switch (format) {
    case ExportFormat.xlsx:
      return (
        bytes: buildXlsx(newestFirst),
        name: 'bang_luong_$stamp.xlsx',
        mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );
    case ExportFormat.pdf:
      return (
        bytes: await buildPdf(
          newestFirst,
          tr('Bảng lương – ${tr(_rangeLabels[range]!)}'),
          rangeSubtitle(range, newestFirst),
        ),
        name: 'bang_luong_$stamp.pdf',
        mime: 'application/pdf',
      );
    case ExportFormat.csv:
      return (
        bytes: utf8.encode(buildCsv(newestFirst)),
        name: 'bang_luong_$stamp.csv',
        mime: 'text/csv',
      );
  }
}

Future<void> shareBytes(List<int> bytes, String name, String mime) async {
  final dir = await getTemporaryDirectory();
  final file = File(p.join(dir.path, name));
  await file.writeAsBytes(bytes, flush: true);
  // share_plus >= 11: dùng SharePlus.instance.share(ShareParams(files: [...]))
  await Share.shareXFiles([XFile(file.path, mimeType: mime)], subject: tr('Bảng lương'));
}

class ExportSheet extends StatefulWidget {
  const ExportSheet({super.key});

  @override
  State<ExportSheet> createState() => _ExportSheetState();
}

class _ExportSheetState extends State<ExportSheet> {
  ExportRange _range = ExportRange.month;

  void _pick(ExportFormat f) => Navigator.pop(context, (range: _range, format: f));

  Widget _tile(IconData icon, Color color, String title, String subtitle, ExportFormat f) => Card(
        child: ListTile(
          leading: Container(
            padding: const EdgeInsets.all(8),
            decoration: BoxDecoration(color: color.withValues(alpha: 0.12), borderRadius: BorderRadius.circular(12)),
            child: Icon(icon, color: color),
          ),
          title: Text(title, style: const TextStyle(fontWeight: FontWeight.w700)),
          subtitle: Text(subtitle, style: const TextStyle(fontSize: 12)),
          trailing: const Icon(Icons.chevron_right_rounded),
          onTap: () => _pick(f),
        ),
      );

  @override
  Widget build(BuildContext context) => SafeArea(
        child: SingleChildScrollView(
          padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            spacing: 12,
            children: [
              Text('Xuất dữ liệu', style: Theme.of(context).textTheme.titleLarge),
              const Text('Khoảng thời gian', style: TextStyle(color: _muted)),
              Wrap(
                spacing: 8,
                runSpacing: 4,
                children: [
                  for (final r in ExportRange.values)
                    ChoiceChip(
                      label: Text(_rangeLabels[r]!),
                      selected: _range == r,
                      onSelected: (_) => setState(() => _range = r),
                    ),
                ],
              ),
              const SizedBox(height: 4),
              const Text('Định dạng', style: TextStyle(color: _muted)),
              _tile(Icons.table_chart_rounded, Colors.green.shade700, 'Excel (.xlsx)',
                  'Chi tiết từng ca và tổng hợp theo tuần', ExportFormat.xlsx),
              _tile(Icons.picture_as_pdf_rounded, Colors.red.shade700, 'PDF',
                  'Báo cáo để in hoặc gửi (chạy offline)', ExportFormat.pdf),
              _tile(Icons.description_outlined, Colors.blueGrey, 'CSV',
                  'Văn bản thuần, mở được bằng Excel hoặc Google Sheets', ExportFormat.csv),
            ],
          ),
        ),
      );
}

// ════════════════════════════════════════════════════════════════════════════
//  UI
// ════════════════════════════════════════════════════════════════════════════

class SalaryApp extends StatefulWidget {
  const SalaryApp({super.key});
  @override
  State<SalaryApp> createState() => _SalaryAppState();
}

class _SalaryAppState extends State<SalaryApp> {
  final _controller = AppController(AppDatabase.instance);

  @override
  void initState() {
    super.initState();
    _controller.load();
  }

  @override
  void dispose() {
    _controller.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => ListenableBuilder(
        listenable: _controller,
        builder: (context, _) {
          currentLang = _controller.settings.lang;
          return _buildApp();
        },
      );

  Widget _buildApp() => MaterialApp(
        key: ValueKey(currentLang),
        title: tr('Theo dõi lương'),
        debugShowCheckedModeBanner: false,
        theme: ThemeData(
          useMaterial3: true,
          colorScheme: ColorScheme.fromSeed(seedColor: _tealMid),
          scaffoldBackgroundColor: _surfaceBg,
          appBarTheme: const AppBarTheme(
            backgroundColor: _surfaceBg,
            surfaceTintColor: Colors.transparent,
            scrolledUnderElevation: 0,
          ),
          cardTheme: CardThemeData(
            elevation: 0,
            color: Colors.white,
            surfaceTintColor: Colors.transparent,
            margin: EdgeInsets.zero,
            shape: RoundedRectangleBorder(
              borderRadius: BorderRadius.circular(20),
              side: const BorderSide(color: Color(0xFFE6ECF0)),
            ),
          ),
        ),
        home: HomePage(c: _controller),
      );
}

void _toast(BuildContext context, String msg) {
  ScaffoldMessenger.of(context)
    ..hideCurrentSnackBar()
    ..showSnackBar(SnackBar(content: Text(msg), behavior: SnackBarBehavior.floating));
}

Future<bool> confirmDialog(
  BuildContext context, {
  required String title,
  required String message,
  required String action,
}) async {
  final ok = await showDialog<bool>(
    context: context,
    builder: (ctx) => AlertDialog(
      title: Text(title),
      content: Text(message),
      actions: [
        TextButton(onPressed: () => Navigator.pop(ctx, false), child: const Text('Hủy')),
        FilledButton(onPressed: () => Navigator.pop(ctx, true), child: Text(action)),
      ],
    ),
  );
  return ok ?? false;
}

class HomePage extends StatelessWidget {
  const HomePage({super.key, required this.c});
  final AppController c;

  @override
  Widget build(BuildContext context) {
    return ListenableBuilder(
      listenable: c,
      builder: (context, _) {
        if (c.loading) {
          return const Scaffold(body: Center(child: CircularProgressIndicator()));
        }
        if (c.loadError != null && c.records.isEmpty) {
          return Scaffold(
            appBar: AppBar(title: const Text('Theo dõi lương')),
            body: Center(
              child: Padding(
                padding: const EdgeInsets.all(24),
                child: Column(
                  mainAxisSize: MainAxisSize.min,
                  spacing: 12,
                  children: [
                    const Icon(Icons.error_outline, size: 40),
                    Text('Không thể tải dữ liệu\n${c.loadError}', textAlign: TextAlign.center),
                    FilledButton(onPressed: c.load, child: const Text('Thử lại')),
                  ],
                ),
              ),
            ),
          );
        }

        final records = c.records;
        final holidays = c.settings.holidays;
        final now = DateTime.now();
        final monday = startOfWeek(now, DateTime.monday);
        final week = summarize(records, monday, DateTime(monday.year, monday.month, monday.day + 7));
        final today = summarize(records, dateOnly(now), DateTime(now.year, now.month, now.day + 1));
        final month = summarize(records, DateTime(now.year, now.month), DateTime(now.year, now.month + 1));
        final year = summarize(records, DateTime(now.year), DateTime(now.year + 1));

        // Gộp lịch sử theo tuần: DateTime = tiêu đề tuần, WorkRecord = ca làm.
        final items = <Object>[];
        DateTime? currentWeek;
        for (final r in records) {
          final w = startOfWeek(r.start, DateTime.monday);
          if (currentWeek != w) {
            currentWeek = w;
            items.add(w);
          }
          items.add(r);
        }

        return Scaffold(
          floatingActionButton: FloatingActionButton.extended(
            onPressed: () => _editRecord(context),
            icon: const Icon(Icons.add_rounded),
            label: const Text('Thêm ca'),
          ),
          body: RefreshIndicator(
            onRefresh: c.refresh,
            child: CustomScrollView(
              physics: const AlwaysScrollableScrollPhysics(),
              slivers: [
                SliverAppBar.large(
                  title: const Text('Theo dõi lương', style: TextStyle(fontWeight: FontWeight.w700)),
                  actions: [
                    IconButton(
                      tooltip: tr('Xuất dữ liệu'),
                      onPressed: () => _export(context),
                      icon: const Icon(Icons.ios_share_rounded),
                    ),
                    IconButton(
                      tooltip: tr('Cài đặt'),
                      onPressed: () => _openSettings(context),
                      icon: const Icon(Icons.tune_rounded),
                    ),
                    const SizedBox(width: 4),
                  ],
                ),
                SliverPadding(
                  padding: const EdgeInsets.fromLTRB(16, 4, 16, 0),
                  sliver: SliverToBoxAdapter(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.stretch,
                      spacing: 14,
                      children: [
                        if (c.skipped > 0)
                          Card(
                            color: Colors.amber.shade100,
                            child: Padding(
                              padding: const EdgeInsets.all(12),
                              child: Text(
                                'Có ${c.skipped} bản ghi cũ không đọc được nên chưa hiển thị. '
                                'Dữ liệu chưa bị xóa và vẫn nằm trong file DB '
                                '(có bản sao work_records_v3.db.pre_v5.bak).',
                                style: const TextStyle(fontSize: 13),
                              ),
                            ),
                          ),
                        ShiftCard(
                          active: c.active,
                          todayType: c.dayTypeFor(now),
                          weekEarned: week.earned,
                          weekHours: week.total,
                          weekShifts: week.shifts,
                          estimate: () => c.active == null
                              ? 0.0
                              : estimateActive(records, c.active!, holidays, DateTime.now()),
                          onStart: () => _startShift(context),
                          onEnd: () => _endShift(context),
                          onDiscard: () => _discardShift(context),
                        ),
                        WeekOverview(records: records, holidays: holidays),
                        Row(
                          spacing: 10,
                          children: [
                            Expanded(
                              child: StatTile(
                                icon: Icons.today_rounded,
                                label: 'Hôm nay',
                                value: money(today.earned),
                                hours: fmtHours(today.total),
                                color: _tealMid,
                              ),
                            ),
                            Expanded(
                              child: StatTile(
                                icon: Icons.calendar_month_rounded,
                                label: 'Tháng này',
                                value: money(month.earned),
                                hours: fmtHours(month.total),
                                color: Colors.indigo,
                              ),
                            ),
                            Expanded(
                              child: StatTile(
                                icon: Icons.emoji_events_outlined,
                                label: 'Năm nay',
                                value: money(year.earned),
                                hours: fmtHours(year.total),
                                color: Colors.deepOrange,
                              ),
                            ),
                          ],
                        ),
                        const SizedBox(height: 4),
                        Row(
                          children: [
                            const Expanded(
                              child: Text(
                                'Lịch sử ca làm',
                                style: TextStyle(fontSize: 20, fontWeight: FontWeight.w700),
                              ),
                            ),
                            Text('${records.length} ca', style: const TextStyle(color: _muted)),
                          ],
                        ),
                      ],
                    ),
                  ),
                ),
                if (records.isEmpty)
                  SliverToBoxAdapter(child: _EmptyState(onAdd: () => _editRecord(context)))
                else
                  SliverPadding(
                    padding: const EdgeInsets.symmetric(horizontal: 16),
                    sliver: SliverList.builder(
                      itemCount: items.length,
                      itemBuilder: (context, i) {
                        final it = items[i];
                        if (it is DateTime) {
                          return WeekHeader(
                            monday: it,
                            summary: summarize(records, it, DateTime(it.year, it.month, it.day + 7)),
                          );
                        }
                        final r = it as WorkRecord;
                        return RecordTile(
                          key: ValueKey(r.id),
                          record: r,
                          onTap: () => _editRecord(context, r),
                          onDelete: () => _deleteRecord(context, r),
                        );
                      },
                    ),
                  ),
                const SliverToBoxAdapter(
                  child: Padding(
                    padding: EdgeInsets.fromLTRB(20, 16, 20, 100),
                    child: Text(
                      'Thu nhập chỉ là ước tính. Mỗi ca được ghi theo ngày bắt đầu '
                      '(kể cả ca qua nửa đêm). Thứ 2–6: 8 giờ đầu x1, 8–12 giờ x1.5, trên 12 giờ x2. '
                      'Thứ 7, CN: 8 giờ đầu được bù vào giờ thường còn thiếu của tuần (tối đa 40 giờ), '
                      'sau đó x1.5, trên 12 giờ x2. Ngày lễ: x2, không tính vào giờ thường; tuần có '
                      'ngày lễ rơi vào thứ 2–6 thì cuối tuần không được bù.',
                      style: TextStyle(fontSize: 12, color: _muted, height: 1.4),
                    ),
                  ),
                ),
              ],
            ),
          ),
        );
      },
    );
  }

  Future<void> _startShift(BuildContext context) async {
    final rate = await showDialog<double>(
      context: context,
      builder: (_) => StartShiftDialog(
        defaultRate: c.settings.hourlyRate,
        dayType: c.dayTypeFor(DateTime.now()),
      ),
    );
    if (rate != null) await c.checkIn(rate);
  }

  Future<void> _endShift(BuildContext context) async {
    final start = c.active?.start;
    if (start == null) return;
    final longShift = DateTime.now().difference(start).inHours >= 16;
    final err = await c.checkOut();
    if (!context.mounted) return;
    _toast(
      context,
      err ??
          (longShift
              ? 'Đã lưu. Ca dài bất thường, hãy kiểm tra lại giờ ra (chạm vào ca để sửa).'
              : 'Đã lưu ca làm.'),
    );
  }

  Future<void> _discardShift(BuildContext context) async {
    final ok = await confirmDialog(
      context,
      title: 'Hủy ca đang làm?',
      message: 'Ca này sẽ không được lưu.',
      action: 'Hủy ca',
    );
    if (ok) await c.discardActive();
  }

  Future<void> _editRecord(BuildContext context, [WorkRecord? existing]) async {
    final draft = await showModalBottomSheet<RecordDraft>(
      context: context,
      isScrollControlled: true,
      useSafeArea: true,
      showDragHandle: true,
      builder: (_) => RecordEditor(
        existing: existing,
        defaultRate: c.settings.hourlyRate,
        validate: c.validate,
        typeOf: c.dayTypeFor,
      ),
    );
    if (draft == null) return;
    final err = await c.saveDraft(draft);
    if (context.mounted) _toast(context, err ?? 'Đã lưu.');
  }

  Future<void> _deleteRecord(BuildContext context, WorkRecord r) async {
    final ok = await confirmDialog(
      context,
      title: 'Xóa ca làm?',
      message: 'Xóa ca ${_dmyFmt.format(r.start)} '
          '(${_timeFmt.format(r.start)}–${_timeFmt.format(r.end)})?',
      action: 'Xóa',
    );
    if (!ok) return;
    final err = await c.delete(r);
    if (context.mounted) _toast(context, err ?? 'Đã xóa.');
  }

  Future<void> _openSettings(BuildContext context) async {
    final s = await showDialog<AppSettings>(
      context: context,
      builder: (_) => SettingsDialog(initial: c.settings),
    );
    if (s != null) await c.updateSettings(s);
  }

  Future<void> _export(BuildContext context) async {
    if (c.records.isEmpty) {
      _toast(context, 'Chưa có dữ liệu để xuất.');
      return;
    }
    final choice = await showModalBottomSheet<({ExportRange range, ExportFormat format})>(
      context: context,
      showDragHandle: true,
      isScrollControlled: true,
      builder: (_) => const ExportSheet(),
    );
    if (choice == null || !context.mounted) return;

    final rs = recordsInRange(c.records, choice.range);
    if (rs.isEmpty) {
      _toast(context, 'Không có ca làm nào trong khoảng này.');
      return;
    }

    // Vòng quay chờ trong lúc tạo file (PDF cần tải font).
    final nav = Navigator.of(context);
    unawaited(showDialog<void>(
      context: context,
      barrierDismissible: false,
      builder: (_) => const PopScope(canPop: false, child: Center(child: CircularProgressIndicator())),
    ));
    ({List<int> bytes, String name, String mime})? file;
    String? error;
    try {
      file = await buildExport(rs, choice.range, choice.format);
    } catch (e) {
      error = '$e';
    }
    nav.pop();

    if (file == null) {
      if (context.mounted) {
        _toast(
          context,
          choice.format == ExportFormat.pdf
              ? 'Không thể tạo PDF (kiểm tra đã khai báo font trong pubspec.yaml): $error'
              : 'Không thể tạo file: $error',
        );
      }
      return;
    }
    try {
      await shareBytes(file.bytes, file.name, file.mime);
    } catch (e) {
      if (context.mounted) _toast(context, 'Không thể chia sẻ file: $e');
    }
  }
}

// ── Thẻ chấm công (hero) ───────────────────────────────────────────────────

class _GlassPill extends StatelessWidget {
  const _GlassPill(this.text);
  final String text;

  @override
  Widget build(BuildContext context) => DecoratedBox(
        decoration: BoxDecoration(
          color: Colors.white.withValues(alpha: 0.18),
          borderRadius: BorderRadius.circular(20),
        ),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 4),
          child: Text(text, style: const TextStyle(color: Colors.white, fontSize: 12, fontWeight: FontWeight.w600)),
        ),
      );
}

class ShiftCard extends StatelessWidget {
  const ShiftCard({
    super.key,
    required this.active,
    required this.todayType,
    required this.weekEarned,
    required this.weekHours,
    required this.weekShifts,
    required this.estimate,
    required this.onStart,
    required this.onEnd,
    required this.onDiscard,
  });
  final ActiveShift? active;
  final String todayType;
  final double weekEarned, weekHours;
  final int weekShifts;
  final double Function() estimate;
  final VoidCallback onStart, onEnd, onDiscard;

  @override
  Widget build(BuildContext context) {
    final a = active;
    final soft = Colors.white.withValues(alpha: 0.78);
    final solid = FilledButton.styleFrom(
      backgroundColor: Colors.white,
      foregroundColor: _tealMid,
      padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 24),
      textStyle: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700),
    );

    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        borderRadius: BorderRadius.circular(28),
        gradient: const LinearGradient(
          begin: Alignment.topLeft,
          end: Alignment.bottomRight,
          colors: [_tealMid, _tealDark],
        ),
        boxShadow: [
          BoxShadow(
            color: _tealMid.withValues(alpha: 0.35),
            blurRadius: 24,
            offset: const Offset(0, 10),
          ),
        ],
      ),
      child: a == null
          ? Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Expanded(
                      child: Text(
                        '${_weekdayFull[DateTime.now().weekday - 1]}, ${_dmyFmt.format(DateTime.now())}',
                        style: TextStyle(color: soft, fontSize: 13),
                      ),
                    ),
                    _GlassPill('Hôm nay: ${dayLabel(todayType)}'),
                  ],
                ),
                const SizedBox(height: 18),
                Text('Thu nhập tuần này', style: TextStyle(color: soft)),
                const SizedBox(height: 2),
                Text(
                  money(weekEarned),
                  style: const TextStyle(
                    color: Colors.white,
                    fontSize: 40,
                    fontWeight: FontWeight.w700,
                    letterSpacing: -0.5,
                  ),
                ),
                Text('${fmtHours(weekHours)} · $weekShifts ca', style: TextStyle(color: soft)),
                const SizedBox(height: 18),
                SizedBox(
                  width: double.infinity,
                  child: FilledButton.icon(
                    style: solid,
                    onPressed: onStart,
                    icon: const Icon(Icons.play_arrow_rounded),
                    label: const Text('Vào ca'),
                  ),
                ),
              ],
            )
          : Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Row(
                  children: [
                    Container(
                      width: 10,
                      height: 10,
                      decoration: const BoxDecoration(color: Color(0xFF86EFAC), shape: BoxShape.circle),
                    ),
                    const SizedBox(width: 8),
                    const Expanded(
                      child: Text(
                        'ĐANG TRONG CA',
                        style: TextStyle(color: Colors.white, fontWeight: FontWeight.w700, letterSpacing: 1),
                      ),
                    ),
                    _GlassPill(dayLabel(a.dayType)),
                  ],
                ),
                const SizedBox(height: 12),
                ElapsedText(
                  start: a.start,
                  style: const TextStyle(color: Colors.white, fontSize: 48, fontWeight: FontWeight.w300),
                ),
                _Tick(
                  builder: (_) => Text(
                    '≈ ${money(estimate())} tính đến lúc này',
                    style: const TextStyle(color: Colors.white, fontSize: 17, fontWeight: FontWeight.w600),
                  ),
                ),
                const SizedBox(height: 4),
                Text(
                  'Vào lúc ${_fmtStart(a.start)} · ${money(a.rate)}/giờ',
                  style: TextStyle(color: soft, fontSize: 13),
                ),
                const SizedBox(height: 18),
                Row(
                  spacing: 10,
                  children: [
                    OutlinedButton(
                      style: OutlinedButton.styleFrom(
                        foregroundColor: Colors.white,
                        side: BorderSide(color: Colors.white.withValues(alpha: 0.6)),
                        padding: const EdgeInsets.symmetric(vertical: 16, horizontal: 18),
                      ),
                      onPressed: onDiscard,
                      child: const Text('Hủy ca'),
                    ),
                    Expanded(
                      child: FilledButton.icon(
                        style: solid,
                        onPressed: onEnd,
                        icon: const Icon(Icons.stop_rounded),
                        label: const Text('Ra ca'),
                      ),
                    ),
                  ],
                ),
              ],
            ),
    );
  }
}

/// Rebuild con của nó mỗi giây (không ảnh hưởng phần còn lại của màn hình).
class _Tick extends StatefulWidget {
  const _Tick({required this.builder});
  final WidgetBuilder builder;

  @override
  State<_Tick> createState() => _TickState();
}

class _TickState extends State<_Tick> {
  Timer? _timer;

  @override
  void initState() {
    super.initState();
    _timer = Timer.periodic(const Duration(seconds: 1), (_) {
      if (mounted) setState(() {});
    });
  }

  @override
  void dispose() {
    _timer?.cancel();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => widget.builder(context);
}

class ElapsedText extends StatelessWidget {
  const ElapsedText({super.key, required this.start, this.style});
  final DateTime start;
  final TextStyle? style;

  @override
  Widget build(BuildContext context) => _Tick(
        builder: (_) {
          var d = DateTime.now().difference(start);
          if (d.isNegative) d = Duration.zero;
          return Text(
            '${_two(d.inHours)}:${_two(d.inMinutes % 60)}:${_two(d.inSeconds % 60)}',
            style: (style ?? const TextStyle())
                .copyWith(fontFeatures: const [FontFeature.tabularFigures()]),
          );
        },
      );
}

// ── Tổng quan tuần ──────────────────────────────────────────────────────────

String _shortH(double h) {
  final m = (h * 60).round();
  if (m < 60) return '${m}p';
  final hh = m / 60;
  return hh == hh.roundToDouble() ? '${hh.round()}h' : '${hh.toStringAsFixed(1)}h';
}

int _flex(double h) => math.max(1, (h * 60).round());

class _Legend extends StatelessWidget {
  const _Legend(this.color, this.text);
  final Color color;
  final String text;

  @override
  Widget build(BuildContext context) => Row(
        mainAxisSize: MainAxisSize.min,
        spacing: 6,
        children: [
          Container(width: 10, height: 10, decoration: BoxDecoration(color: color, shape: BoxShape.circle)),
          Text(text, style: const TextStyle(fontSize: 12, color: _muted)),
        ],
      );
}

class _DayBar extends StatelessWidget {
  const _DayBar({
    required this.label,
    required this.today,
    required this.holiday,
    required this.reg,
    required this.ot,
    required this.dbl,
    required this.maxH,
  });
  final String label;
  final bool today, holiday;
  final double reg, ot, dbl, maxH;

  @override
  Widget build(BuildContext context) {
    final total = reg + ot + dbl;
    const barArea = 78.0;
    final h = total <= 0 ? 4.0 : math.max(6.0, barArea * total / maxH);

    return Container(
      padding: const EdgeInsets.symmetric(vertical: 6),
      decoration: BoxDecoration(
        color: today ? _tealMid.withValues(alpha: 0.08) : null,
        borderRadius: BorderRadius.circular(14),
      ),
      child: Column(
        children: [
          SizedBox(
            height: 14,
            child: total > 0
                ? Text(_shortH(total),
                    style: const TextStyle(fontSize: 10, color: _muted, fontWeight: FontWeight.w600))
                : null,
          ),
          const SizedBox(height: 2),
          SizedBox(
            height: barArea,
            child: Align(
              alignment: Alignment.bottomCenter,
              child: AnimatedContainer(
                duration: const Duration(milliseconds: 500),
                curve: Curves.easeOutCubic,
                width: 18,
                height: h,
                clipBehavior: Clip.antiAlias,
                decoration: BoxDecoration(
                  color: total <= 0 ? const Color(0xFFE2E8F0) : null,
                  borderRadius: BorderRadius.circular(6),
                ),
                child: total <= 0
                    ? null
                    : Column(
                        children: [
                          if (dbl > 0) Expanded(flex: _flex(dbl), child: const ColoredBox(color: Colors.deepOrange)),
                          if (ot > 0) Expanded(flex: _flex(ot), child: const ColoredBox(color: Colors.indigo)),
                          if (reg > 0) Expanded(flex: _flex(reg), child: const ColoredBox(color: _tealMid)),
                        ],
                      ),
              ),
            ),
          ),
          const SizedBox(height: 6),
          Text(
            holiday ? '$label★' : label,
            style: TextStyle(
              fontSize: 12,
              fontWeight: today ? FontWeight.w700 : FontWeight.w500,
              color: holiday ? Colors.deepOrange : (today ? _tealMid : _muted),
            ),
          ),
        ],
      ),
    );
  }
}

class WeekOverview extends StatelessWidget {
  const WeekOverview({super.key, required this.records, required this.holidays});
  final List<WorkRecord> records;
  final Set<String> holidays;

  @override
  Widget build(BuildContext context) {
    final now = DateTime.now();
    final monday = startOfWeek(now, DateTime.monday);
    final sunday = DateTime(monday.year, monday.month, monday.day + 6);
    final next = DateTime(monday.year, monday.month, monday.day + 7);
    final days = [for (var i = 0; i < 7; i++) DateTime(monday.year, monday.month, monday.day + i)];

    final reg = List.filled(7, 0.0), ot = List.filled(7, 0.0), dbl = List.filled(7, 0.0);
    for (final r in records) {
      if (r.start.isBefore(monday) || !r.start.isBefore(next)) continue;
      final i = r.start.weekday - 1;
      reg[i] += r.regularHours;
      ot[i] += r.otHours;
      dbl[i] += r.doubleHours;
    }
    final maxH = math.max(
      10.0,
      [for (var i = 0; i < 7; i++) reg[i] + ot[i] + dbl[i]].fold<double>(0, (a, b) => math.max(a, b)),
    );
    final weekReg = reg.fold<double>(0, (a, b) => a + b);
    final missing = math.max(0.0, 40 - weekReg);
    final noFill = PayCalculator._weekdayHolidayIn(monday, holidays);

    final String note;
    if (missing <= 0) {
      note = 'Đã đủ 40 giờ thường. Giờ làm thêm tính OT.';
    } else if (noFill) {
      note = 'Còn thiếu ${fmtHours(missing)}. Tuần có ngày lễ T2–T6 nên cuối tuần không được bù.';
    } else {
      note = 'Còn thiếu ${fmtHours(missing)}. Thứ 7, CN có thể bù tối đa 8 giờ mỗi ngày.';
    }

    return Card(
      child: Padding(
        padding: const EdgeInsets.all(18),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              children: [
                const Expanded(
                  child: Text('Tuần này', style: TextStyle(fontSize: 17, fontWeight: FontWeight.w700)),
                ),
                Text(
                  '${_two(monday.day)}/${_two(monday.month)} – ${_two(sunday.day)}/${_two(sunday.month)}',
                  style: const TextStyle(color: _muted),
                ),
              ],
            ),
            const SizedBox(height: 12),
            Row(
              spacing: 4,
              children: [
                for (var i = 0; i < 7; i++)
                  Expanded(
                    child: _DayBar(
                      label: _weekdayShort[i],
                      today: dateOnly(now) == days[i],
                      holiday: holidays.contains(_dateFmt.format(days[i])),
                      reg: reg[i],
                      ot: ot[i],
                      dbl: dbl[i],
                      maxH: maxH,
                    ),
                  ),
              ],
            ),
            const SizedBox(height: 12),
            const Wrap(
              spacing: 14,
              runSpacing: 4,
              children: [
                _Legend(_tealMid, 'Thường'),
                _Legend(Colors.indigo, 'OT x1.5'),
                _Legend(Colors.deepOrange, 'x2'),
              ],
            ),
            const SizedBox(height: 16),
            Row(
              children: [
                const Expanded(child: Text('Giờ thường trong tuần', style: TextStyle(color: _muted))),
                Text('${fmtHours(weekReg)} / 40h', style: const TextStyle(fontWeight: FontWeight.w700)),
              ],
            ),
            const SizedBox(height: 8),
            ClipRRect(
              borderRadius: BorderRadius.circular(8),
              child: TweenAnimationBuilder<double>(
                tween: Tween<double>(begin: 0, end: (weekReg / 40).clamp(0.0, 1.0)),
                duration: const Duration(milliseconds: 600),
                curve: Curves.easeOutCubic,
                builder: (_, v, __) => LinearProgressIndicator(
                  value: v,
                  minHeight: 10,
                  backgroundColor: const Color(0xFFE2E8F0),
                  color: _tealMid,
                ),
              ),
            ),
            const SizedBox(height: 8),
            Text(note, style: const TextStyle(fontSize: 12, color: _muted, height: 1.35)),
          ],
        ),
      ),
    );
  }
}

class StatTile extends StatelessWidget {
  const StatTile({
    super.key,
    required this.icon,
    required this.label,
    required this.value,
    required this.hours,
    required this.color,
  });
  final IconData icon;
  final String label, value, hours;
  final Color color;

  @override
  Widget build(BuildContext context) => Card(
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Container(
                padding: const EdgeInsets.all(7),
                decoration: BoxDecoration(
                  color: color.withValues(alpha: 0.12),
                  borderRadius: BorderRadius.circular(10),
                ),
                child: Icon(icon, size: 18, color: color),
              ),
              const SizedBox(height: 10),
              Text(label, style: const TextStyle(fontSize: 12, color: _muted)),
              const SizedBox(height: 2),
              FittedBox(
                fit: BoxFit.scaleDown,
                alignment: Alignment.centerLeft,
                child: Text(value, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.w700)),
              ),
              Text(hours, style: const TextStyle(fontSize: 12, color: _muted)),
            ],
          ),
        ),
      );
}

// ── Lịch sử ─────────────────────────────────────────────────────────────────

class WeekHeader extends StatelessWidget {
  const WeekHeader({super.key, required this.monday, required this.summary});
  final DateTime monday;
  final PeriodSummary summary;

  @override
  Widget build(BuildContext context) {
    final sunday = DateTime(monday.year, monday.month, monday.day + 6);
    return Padding(
      padding: const EdgeInsets.fromLTRB(4, 16, 4, 8),
      child: Row(
        children: [
          Expanded(
            child: Column(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Text(
                  'Tuần ${_two(monday.day)}/${_two(monday.month)} – ${_two(sunday.day)}/${_two(sunday.month)}',
                  style: const TextStyle(fontWeight: FontWeight.w700),
                ),
                Text(
                  'Thường ${fmtHours(summary.regular)} · OT ${fmtHours(summary.ot)} · x2 ${fmtHours(summary.dbl)}',
                  style: const TextStyle(fontSize: 12, color: _muted),
                ),
              ],
            ),
          ),
          Text(
            money(summary.earned),
            style: const TextStyle(fontWeight: FontWeight.w700, color: _tealMid, fontSize: 15),
          ),
        ],
      ),
    );
  }
}

class RecordTile extends StatelessWidget {
  const RecordTile({
    super.key,
    required this.record,
    required this.onTap,
    required this.onDelete,
  });
  final WorkRecord record;
  final VoidCallback onTap, onDelete;

  @override
  Widget build(BuildContext context) {
    final r = record;
    final color = typeColor(r.dayType);
    final range = '${_timeFmt.format(r.start)} – ${_timeFmt.format(r.end)}${r.overnight ? ' (+1)' : ''}';

    return Padding(
      padding: const EdgeInsets.only(bottom: 8),
      child: Card(
        clipBehavior: Clip.antiAlias,
        child: InkWell(
          onTap: onTap,
          child: Padding(
            padding: const EdgeInsets.fromLTRB(12, 12, 4, 12),
            child: Row(
              children: [
                Container(
                  width: 50,
                  height: 56,
                  decoration: BoxDecoration(
                    color: color.withValues(alpha: 0.12),
                    borderRadius: BorderRadius.circular(14),
                  ),
                  child: Column(
                    mainAxisAlignment: MainAxisAlignment.center,
                    children: [
                      Text(
                        _weekdayShort[r.start.weekday - 1],
                        style: TextStyle(fontSize: 11, color: color, fontWeight: FontWeight.w600),
                      ),
                      Text(
                        '${r.start.day}',
                        style: TextStyle(fontSize: 20, color: color, fontWeight: FontWeight.w700),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Flexible(
                            child: Text(range, style: const TextStyle(fontWeight: FontWeight.w700)),
                          ),
                          if (r.dayType != 'Thường') ...[
                            const SizedBox(width: 6),
                            _Pill(dayLabel(r.dayType), color: color),
                          ],
                        ],
                      ),
                      const SizedBox(height: 6),
                      Wrap(
                        spacing: 6,
                        runSpacing: 4,
                        children: [
                          if (r.regularHours > 0) _Pill('Thường ${fmtHours(r.regularHours)}', color: _tealMid),
                          if (r.otHours > 0) _Pill('OT ${fmtHours(r.otHours)}', color: Colors.indigo),
                          if (r.doubleHours > 0) _Pill('x2 ${fmtHours(r.doubleHours)}', color: Colors.deepOrange),
                          _Pill('${money(r.hourlyRate)}/giờ', color: Colors.blueGrey),
                        ],
                      ),
                    ],
                  ),
                ),
                Column(
                  crossAxisAlignment: CrossAxisAlignment.end,
                  children: [
                    Text(money(r.earned), style: const TextStyle(fontWeight: FontWeight.w700, fontSize: 16)),
                    Text(fmtHours(r.totalHours), style: const TextStyle(fontSize: 12, color: _muted)),
                  ],
                ),
                PopupMenuButton<String>(
                  icon: const Icon(Icons.more_vert, size: 20, color: _muted),
                  onSelected: (v) => v == 'edit' ? onTap() : onDelete(),
                  itemBuilder: (_) => const [
                    PopupMenuItem(value: 'edit', child: Text('Sửa')),
                    PopupMenuItem(value: 'delete', child: Text('Xóa')),
                  ],
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}

class _EmptyState extends StatelessWidget {
  const _EmptyState({required this.onAdd});
  final VoidCallback onAdd;

  @override
  Widget build(BuildContext context) => Padding(
        padding: const EdgeInsets.symmetric(vertical: 36, horizontal: 24),
        child: Column(
          spacing: 10,
          children: [
            Container(
              padding: const EdgeInsets.all(18),
              decoration: BoxDecoration(color: _tealMid.withValues(alpha: 0.1), shape: BoxShape.circle),
              child: const Icon(Icons.event_note_rounded, size: 36, color: _tealMid),
            ),
            const Text('Chưa có ca làm nào', style: TextStyle(fontSize: 16, fontWeight: FontWeight.w700)),
            const Text(
              'Bấm “Vào ca” để chấm công hoặc thêm một ca thủ công.',
              textAlign: TextAlign.center,
              style: TextStyle(color: _muted),
            ),
            OutlinedButton.icon(
              onPressed: onAdd,
              icon: const Icon(Icons.add_rounded),
              label: const Text('Thêm ca'),
            ),
          ],
        ),
      );
}

class _Pill extends StatelessWidget {
  const _Pill(this.text, {required this.color});
  final String text;
  final Color color;

  @override
  Widget build(BuildContext context) => DecoratedBox(
        decoration: BoxDecoration(
          color: color.withValues(alpha: 0.1),
          borderRadius: BorderRadius.circular(20),
        ),
        child: Padding(
          padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 3),
          child: Text(text, style: TextStyle(fontSize: 12, color: color, fontWeight: FontWeight.w500)),
        ),
      );
}

// ── Form & hộp thoại ────────────────────────────────────────────────────────

class StartShiftDialog extends StatefulWidget {
  const StartShiftDialog({super.key, required this.defaultRate, required this.dayType});
  final double defaultRate;
  final String dayType;

  @override
  State<StartShiftDialog> createState() => _StartShiftDialogState();
}

class _StartShiftDialogState extends State<StartShiftDialog> {
  late final TextEditingController _rate =
      TextEditingController(text: _rateFmt.format(widget.defaultRate));
  String? _error;

  @override
  void dispose() {
    _rate.dispose();
    super.dispose();
  }

  void _submit() {
    final rate = parseNumber(_rate.text);
    if (rate == null) {
      setState(() => _error = 'Mức lương không hợp lệ.');
      return;
    }
    Navigator.pop(context, rate);
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Bắt đầu ca làm'),
        content: SingleChildScrollView(
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            spacing: 16,
            children: [
              _Pill('Hôm nay tính là: ${dayLabel(widget.dayType)}', color: typeColor(widget.dayType)),
              TextField(
                controller: _rate,
                keyboardType: const TextInputType.numberWithOptions(decimal: true),
                decoration: InputDecoration(
                  labelText: tr('Lương theo giờ'),
                  prefixText: r'$ ',
                  errorText: _error == null ? null : tr(_error!),
                  border: const OutlineInputBorder(),
                ),
              ),
            ],
          ),
        ),
        actions: [
          TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
          FilledButton(onPressed: _submit, child: const Text('Vào ca')),
        ],
      );
}

class SettingsDialog extends StatefulWidget {
  const SettingsDialog({super.key, required this.initial});
  final AppSettings initial;

  @override
  State<SettingsDialog> createState() => _SettingsDialogState();
}

class _SettingsDialogState extends State<SettingsDialog> {
  late final TextEditingController _rate =
      TextEditingController(text: _rateFmt.format(widget.initial.hourlyRate));
  late final Set<String> _holidays = {...widget.initial.holidays};
  late String _lang = widget.initial.lang;
  String? _error;

  @override
  void dispose() {
    _rate.dispose();
    super.dispose();
  }

  Future<void> _addHoliday() async {
    final d = await showDatePicker(
      context: context,
      initialDate: DateTime.now(),
      firstDate: DateTime(2020),
      lastDate: DateTime(2100),
      helpText: tr('Chọn ngày lễ'),
    );
    if (d != null) setState(() => _holidays.add(_dateFmt.format(d)));
  }

  String _label(String key) {
    final d = DateTime.tryParse(key);
    return d == null ? key : '${_weekdayShort[d.weekday - 1]}, ${_dmyFmt.format(d)}';
  }

  void _submit() {
    final rate = parseNumber(_rate.text);
    if (rate == null) {
      setState(() => _error = 'Mức lương không hợp lệ.');
      return;
    }
    Navigator.pop(context, AppSettings(hourlyRate: rate, holidays: _holidays, lang: _lang));
  }

  @override
  Widget build(BuildContext context) {
    final sorted = _holidays.toList()..sort();
    return AlertDialog(
      title: const Text('Cài đặt'),
      content: SingleChildScrollView(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          spacing: 14,
          children: [
            const Text('Ngôn ngữ / Language', style: TextStyle(fontWeight: FontWeight.w600)),
            SegmentedButton<String>(
              showSelectedIcon: false,
              segments: [
                for (final e in supportedLangs.entries) ButtonSegment(value: e.key, label: Text(e.value)),
              ],
              selected: {_lang},
              onSelectionChanged: (v) => setState(() => _lang = v.first),
            ),
            TextField(
              controller: _rate,
              keyboardType: const TextInputType.numberWithOptions(decimal: true),
              decoration: InputDecoration(
                labelText: tr('Lương mặc định (mỗi giờ)'),
                prefixText: r'$ ',
                errorText: _error == null ? null : tr(_error!),
                border: const OutlineInputBorder(),
              ),
            ),
            const Text('Ngày lễ (tính x2)', style: TextStyle(fontWeight: FontWeight.w600)),
            if (sorted.isEmpty)
              const Text('Chưa đánh dấu ngày lễ nào.',
                  style: TextStyle(color: Colors.black54))
            else
              Wrap(
                spacing: 8,
                runSpacing: 4,
                children: [
                  for (final h in sorted)
                    InputChip(
                      label: Text(_label(h)),
                      onDeleted: () => setState(() => _holidays.remove(h)),
                    ),
                ],
              ),
            OutlinedButton.icon(
              onPressed: _addHoliday,
              icon: const Icon(Icons.event_available),
              label: const Text('Thêm ngày lễ'),
            ),
            const Text(
              'Tuần làm việc bắt đầu thứ 2. Thứ 2–6 là ngày thường, thứ 7 và CN là ngày OT '
              '(tự động). Lương mặc định chỉ áp dụng cho ca mới.',
              style: TextStyle(fontSize: 12, color: Colors.black54),
            ),
          ],
        ),
      ),
      actions: [
        TextButton(onPressed: () => Navigator.pop(context), child: const Text('Hủy')),
        FilledButton(onPressed: _submit, child: const Text('Lưu')),
      ],
    );
  }
}

class RecordEditor extends StatefulWidget {
  const RecordEditor({
    super.key,
    required this.existing,
    required this.defaultRate,
    required this.validate,
    required this.typeOf,
  });
  final WorkRecord? existing;
  final double defaultRate;
  final String? Function(RecordDraft) validate;
  final String Function(DateTime) typeOf;

  @override
  State<RecordEditor> createState() => _RecordEditorState();
}

class _RecordEditorState extends State<RecordEditor> {
  late DateTime _date;
  late TimeOfDay _in;
  late TimeOfDay _out;
  late final TextEditingController _rate;
  String? _error;

  @override
  void initState() {
    super.initState();
    final e = widget.existing;
    _date = dateOnly(e?.start ?? DateTime.now());
    _in = e != null ? TimeOfDay.fromDateTime(e.start) : const TimeOfDay(hour: 9, minute: 0);
    _out = e != null ? TimeOfDay.fromDateTime(e.end) : const TimeOfDay(hour: 17, minute: 0);
    _rate = TextEditingController(text: _rateFmt.format(e?.hourlyRate ?? widget.defaultRate));
  }

  @override
  void dispose() {
    _rate.dispose();
    super.dispose();
  }

  DateTime get _startDt =>
      DateTime(_date.year, _date.month, _date.day, _in.hour, _in.minute);

  DateTime get _endDt {
    final e = DateTime(_date.year, _date.month, _date.day, _out.hour, _out.minute);
    return e.isAfter(_startDt)
        ? e
        : DateTime(e.year, e.month, e.day + 1, e.hour, e.minute);
  }

  String _tod(TimeOfDay t) => '${_two(t.hour)}:${_two(t.minute)}';

  Future<void> _pickDate() async {
    final d = await showDatePicker(
      context: context,
      initialDate: _date,
      firstDate: DateTime(2020),
      lastDate: DateTime(2100),
    );
    if (d != null) setState(() {
      _date = dateOnly(d);
      _error = null;
    });
  }

  Future<void> _pickTime(bool isIn) async {
    final t = await showTimePicker(context: context, initialTime: isIn ? _in : _out);
    if (t == null) return;
    setState(() {
      if (isIn) {
        _in = t;
      } else {
        _out = t;
      }
      _error = null;
    });
  }

  void _save() {
    final rate = parseNumber(_rate.text);
    if (rate == null) {
      setState(() => _error = 'Mức lương không hợp lệ.');
      return;
    }
    if (_in == _out) {
      setState(() => _error = 'Giờ ra phải khác giờ vào.');
      return;
    }
    final draft = RecordDraft(
      id: widget.existing?.id,
      start: _startDt,
      end: _endDt,
      rate: rate,
    );
    final err = widget.validate(draft);
    if (err != null) {
      setState(() => _error = err);
      return;
    }
    Navigator.pop(context, draft);
  }

  @override
  Widget build(BuildContext context) {
    final cs = Theme.of(context).colorScheme;
    final minutes = _in == _out ? 0 : _endDt.difference(_startDt).inMinutes;
    final overnight = dateOnly(_endDt) != dateOnly(_startDt);

    return Padding(
      padding: EdgeInsets.fromLTRB(20, 0, 20, 20 + MediaQuery.viewInsetsOf(context).bottom),
      child: SingleChildScrollView(
        child: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.stretch,
          spacing: 14,
          children: [
            Text(
              widget.existing == null ? 'Thêm ca làm' : 'Sửa ca làm',
              style: Theme.of(context).textTheme.titleLarge,
            ),
            OutlinedButton.icon(
              onPressed: _pickDate,
              icon: const Icon(Icons.event),
              label: Text(
                  '${_weekdayShort[_date.weekday - 1]}, ${_dmyFmt.format(_date)}'),
            ),
            Row(
              spacing: 10,
              children: [
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _pickTime(true),
                    icon: const Icon(Icons.login),
                    label: Text('Vào ${_tod(_in)}'),
                  ),
                ),
                Expanded(
                  child: OutlinedButton.icon(
                    onPressed: () => _pickTime(false),
                    icon: const Icon(Icons.logout),
                    label: Text('Ra ${_tod(_out)}'),
                  ),
                ),
              ],
            ),
            Text(
              'Thời lượng: ${fmtHours(minutes / 60)}${overnight && minutes > 0 ? ' · qua nửa đêm' : ''}',
              style: const TextStyle(color: Colors.black54),
            ),
            Align(
              alignment: Alignment.centerLeft,
              child: _Pill('Loại ngày: ${dayLabel(widget.typeOf(_date))}', color: typeColor(widget.typeOf(_date))),
            ),
            TextField(
              controller: _rate,
              keyboardType: const TextInputType.numberWithOptions(decimal: true),
              decoration: InputDecoration(
                labelText: tr('Lương theo giờ'),
                prefixText: r'$ ',
                border: const OutlineInputBorder(),
              ),
            ),
            if (_error != null) Text(_error!, style: TextStyle(color: cs.error)),
            FilledButton(
              onPressed: _save,
              style: FilledButton.styleFrom(padding: const EdgeInsets.symmetric(vertical: 14)),
              child: const Text('Lưu'),
            ),
          ],
        ),
      ),
    );
  }
}
