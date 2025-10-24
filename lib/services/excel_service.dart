import 'dart:io';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';
import 'package:permission_handler/permission_handler.dart';
import '../models/students.dart';

class ExcelService {
  static String _fileName = 'student_details.xlsx';
  static String _customPath = '';
  static const String sheetName = 'Students';
  static List<Student> _studentBuffer = [];

  // Request storage permission
  static Future<bool> requestStoragePermission() async {
    if (Platform.isAndroid) {
      var status = await Permission.storage.status;
      if (status.isGranted) return true;

      var result = await Permission.storage.request();
      if (result.isGranted) return true;

      if (await Permission.manageExternalStorage.isGranted) return true;

      var manageResult = await Permission.manageExternalStorage.request();
      return manageResult.isGranted;
    }
    return true;
  }

  // Set custom file path and name
  static void setCustomFileInfo(String fileName, String path) {
    _fileName = fileName.endsWith('.xlsx') ? fileName : '$fileName.xlsx';
    _customPath = path;
  }

  // Get the file path
  static Future<String> getFilePath() async {
    if (_customPath.isNotEmpty) {
      return '$_customPath/$_fileName';
    }
    
    Directory? directory;
    if (Platform.isAndroid) {
      directory = Directory('/storage/emulated/0/Download');
      if (!await directory.exists()) {
        directory = await getApplicationDocumentsDirectory();
      }
    } else if (Platform.isIOS) {
      directory = await getApplicationDocumentsDirectory();
    }

    final path = directory!.path;
    return '$path/$_fileName';
  }

  // Check if file exists
  static Future<bool> fileExists() async {
    final filePath = await getFilePath();
    final file = File(filePath);
    return await file.exists();
  }

  // Create new Excel with headers
  static Future<Excel> createNewExcel() async {
    var excel = Excel.createExcel();
    excel.rename(excel.getDefaultSheet()!, sheetName);
    Sheet sheet = excel[sheetName];
    
    sheet.appendRow([
      TextCellValue('Name'),
      TextCellValue('Phone Number'),
    ]);
    
    CellStyle headerStyle = CellStyle(
      bold: true,
      backgroundColorHex: ExcelColor.blue700,
      fontColorHex: ExcelColor.white,
    );
    
    sheet.cell(CellIndex.indexByString('A1')).cellStyle = headerStyle;
    sheet.cell(CellIndex.indexByString('B1')).cellStyle = headerStyle;
    
    return excel;
  }

  // Create Excel at custom path
  static Future<bool> createExcelAtPath(String fileName, String path) async {
    try {
      bool hasPermission = await requestStoragePermission();
      if (!hasPermission) {
        print('Storage permission denied');
        return false;
      }

      setCustomFileInfo(fileName, path);
      Excel excel = await createNewExcel();
      final filePath = await getFilePath();
      final file = File(filePath);

      var fileBytes = excel.encode();
      if (fileBytes != null) {
        await file.writeAsBytes(fileBytes);
        print('Excel file created successfully at: $filePath');
        _studentBuffer.clear();
        return true;
      }
      return false;
    } catch (e) {
      print('Error creating Excel file: $e');
      return false;
    }
  }

  // Load existing Excel
  static Future<Excel?> loadExistingExcel() async {
    try {
      final filePath = await getFilePath();
      final file = File(filePath);
      
      if (await file.exists()) {
        var bytes = await file.readAsBytes();
        var excel = Excel.decodeBytes(bytes);
        return excel;
      }
      return null;
    } catch (e) {
      print('Error loading Excel file: $e');
      return null;
    }
  }

  // Add student to buffer
  static void addStudentToBuffer(Student student) {
    _studentBuffer.add(student);
    print('Student added to buffer. Total: ${_studentBuffer.length}');
  }

  // Save all buffered students to Excel
  static Future<bool> saveBufferToExcel() async {
    try {
      if (_studentBuffer.isEmpty) {
        print('No students in buffer to save');
        return true;
      }

      bool hasPermission = await requestStoragePermission();
      if (!hasPermission) {
        print('Storage permission denied');
        return false;
      }

      Excel excel;
      if (await fileExists()) {
        Excel? existingExcel = await loadExistingExcel();
        excel = existingExcel ?? await createNewExcel();
      } else {
        excel = await createNewExcel();
      }

      Sheet sheet = excel[sheetName];

      for (var student in _studentBuffer) {
        sheet.appendRow([
          TextCellValue(student.name),
          TextCellValue(student.phoneNumber),
        ]);
      }

      final filePath = await getFilePath();
      final file = File(filePath);
      var fileBytes = excel.encode();
      
      if (fileBytes != null) {
        await file.writeAsBytes(fileBytes);
        print('${_studentBuffer.length} students saved at: $filePath');
        _studentBuffer.clear();
        return true;
      }
      return false;
    } catch (e) {
      print('Error saving students: $e');
      return false;
    }
  }

  // Get buffer count
  static int getBufferCount() => _studentBuffer.length;

  // Clear buffer
  static void clearBuffer() => _studentBuffer.clear();

  // Get all students from Excel
  static Future<List<Student>> getAllStudents() async {
    try {
      Excel? excel = await loadExistingExcel();
      if (excel == null) return [];

      Sheet sheet = excel[sheetName];
      List<Student> students = [];

      for (int i = 1; i < sheet.maxRows; i++) {
        var row = sheet.row(i);
        if (row.length >= 2) {
          var nameCell = row[0];
          var phoneCell = row[1];

          if (nameCell != null && phoneCell != null) {
            students.add(Student(
              name: nameCell.value.toString(),
              phoneNumber: phoneCell.value.toString(),
            ));
          }
        }
      }
      return students;
    } catch (e) {
      print('Error reading students: $e');
      return [];
    }
  }

  // Download to original path
  static Future<bool> downloadToOriginalPath() async {
    try {
      bool hasPermission = await requestStoragePermission();
      if (!hasPermission) {
        print('Storage permission denied');
        return false;
      }

      final currentPath = await getFilePath();
      final currentFile = File(currentPath);

      if (!await currentFile.exists()) {
        print('Source file does not exist');
        return false;
      }

      print('File already at destination: $currentPath');
      return true;
    } catch (e) {
      print('Error in download: $e');
      return false;
    }
  }

  // Get default download path
  static Future<String> getDefaultDownloadPath() async {
    if (Platform.isAndroid) {
      return '/storage/emulated/0/Download';
    } else if (Platform.isIOS) {
      final directory = await getApplicationDocumentsDirectory();
      return directory.path;
    }
    return '/storage/emulated/0/Download';
  }

  // Validate path
  static Future<bool> isPathValid(String path) async {
    try {
      final directory = Directory(path);
      if (!await directory.exists()) {
        await directory.create(recursive: true);
      }
      return true;
    } catch (e) {
      print('Invalid path: $e');
      return false;
    }
  }
}
