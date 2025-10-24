import 'package:flutter/material.dart';
import 'package:file_picker/file_picker.dart';
import '../services/excel_service.dart';
import 'home_screen.dart';

class ExcelCreationScreen extends StatefulWidget {
  const ExcelCreationScreen({Key? key}) : super(key: key);

  @override
  State<ExcelCreationScreen> createState() => _ExcelCreationScreenState();
}

class _ExcelCreationScreenState extends State<ExcelCreationScreen> {
  final _formKey = GlobalKey<FormState>();
  final _fileNameController = TextEditingController(text: 'student_details');
  String _selectedPath = '';
  bool _isCreating = false;

  @override
  void initState() {
    super.initState();
    _loadDefaultPath();
  }

  @override
  void dispose() {
    _fileNameController.dispose();
    super.dispose();
  }

  Future<void> _loadDefaultPath() async {
    String defaultPath = await ExcelService.getDefaultDownloadPath();
    setState(() {
      _selectedPath = defaultPath;
    });
  }

  Future<void> _pickDirectory() async {
    try {
      String? selectedDirectory = await FilePicker.platform.getDirectoryPath();
      
      if (selectedDirectory != null) {
        setState(() {
          _selectedPath = selectedDirectory;
        });
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('Error selecting directory: $e'),
            backgroundColor: Colors.red,
          ),
        );
      }
    }
  }

  String? _validateFileName(String? value) {
    if (value == null || value.isEmpty) {
      return 'Please enter a file name';
    }
    
    final invalidChars = RegExp(r'[<>:"/\\|?*]');
    if (invalidChars.hasMatch(value)) {
      return 'File name contains invalid characters';
    }
    
    return null;
  }

  Future<void> _createExcelAndNavigate() async {
    if (_formKey.currentState!.validate()) {
      if (_selectedPath.isEmpty) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text('Please select a path'),
            backgroundColor: Colors.orange,
          ),
        );
        return;
      }

      setState(() {
        _isCreating = true;
      });

      bool pathValid = await ExcelService.isPathValid(_selectedPath);
      
      if (!pathValid) {
        setState(() {
          _isCreating = false;
        });
        if (mounted) {
          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(
              content: Text('Invalid path. Please select a valid directory.'),
              backgroundColor: Colors.red,
            ),
          );
        }
        return;
      }

      String fileName = _fileNameController.text.trim();
      bool success = await ExcelService.createExcelAtPath(fileName, _selectedPath);

      setState(() {
        _isCreating = false;
      });

      if (success) {
        if (mounted) {
          ScaffoldMessenger.of(context).showSnackBar(
            SnackBar(
              content: Text('Excel file created at $_selectedPath/$fileName.xlsx'),
              backgroundColor: Colors.green,
              duration: const Duration(seconds: 3),
            ),
          );

          Navigator.pushReplacement(
            context,
            MaterialPageRoute(
              builder: (context) => const HomeScreen(),
            ),
          );
        }
      } else {
        if (mounted) {
          ScaffoldMessenger.of(context).showSnackBar(
            const SnackBar(
              content: Text('Failed to create Excel file. Check permissions.'),
              backgroundColor: Colors.red,
              duration: Duration(seconds: 3),
            ),
          );
        }
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: const Text('Create Excel File'),
        centerTitle: true,
        elevation: 2,
      ),
      body: SingleChildScrollView(
        padding: const EdgeInsets.all(20.0),
        child: Form(
          key: _formKey,
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.stretch,
            children: [
              const SizedBox(height: 20),
              
              const Icon(
                Icons.create_new_folder,
                size: 100,
                color: Colors.blue,
              ),
              
              const SizedBox(height: 30),
              
              const Text(
                'Create New Excel File',
                style: TextStyle(
                  fontSize: 26,
                  fontWeight: FontWeight.bold,
                ),
                textAlign: TextAlign.center,
              ),
              
              const SizedBox(height: 10),
              
              const Text(
                'Choose where to save your student data',
                style: TextStyle(
                  fontSize: 14,
                  color: Colors.grey,
                ),
                textAlign: TextAlign.center,
              ),
              
              const SizedBox(height: 40),
              
              TextFormField(
                controller: _fileNameController,
                decoration: InputDecoration(
                  labelText: 'File Name',
                  hintText: 'Enter file name (without extension)',
                  prefixIcon: const Icon(Icons.description),
                  suffixText: '.xlsx',
                  border: OutlineInputBorder(
                    borderRadius: BorderRadius.circular(10),
                  ),
                  filled: true,
                  fillColor: Colors.grey[100],
                ),
                validator: _validateFileName,
              ),
              
              const SizedBox(height: 20),
              
              Card(
                elevation: 2,
                shape: RoundedRectangleBorder(
                  borderRadius: BorderRadius.circular(10),
                ),
                child: Padding(
                  padding: const EdgeInsets.all(16.0),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          const Icon(Icons.folder_open, color: Colors.blue),
                          const SizedBox(width: 10),
                          const Text(
                            'Save Location',
                            style: TextStyle(
                              fontSize: 16,
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                        ],
                      ),
                      
                      const SizedBox(height: 12),
                      
                      Container(
                        padding: const EdgeInsets.all(12),
                        decoration: BoxDecoration(
                          color: Colors.grey[200],
                          borderRadius: BorderRadius.circular(8),
                        ),
                        child: Text(
                          _selectedPath.isEmpty ? 'No path selected' : _selectedPath,
                          style: TextStyle(
                            fontSize: 13,
                            color: _selectedPath.isEmpty ? Colors.grey : Colors.black87,
                          ),
                        ),
                      ),
                      
                      const SizedBox(height: 12),
                      
                      SizedBox(
                        width: double.infinity,
                        child: OutlinedButton.icon(
                          onPressed: _pickDirectory,
                          icon: const Icon(Icons.folder_special),
                          label: const Text('Choose Location'),
                          style: OutlinedButton.styleFrom(
                            padding: const EdgeInsets.symmetric(vertical: 12),
                            shape: RoundedRectangleBorder(
                              borderRadius: BorderRadius.circular(8),
                            ),
                          ),
                        ),
                      ),
                    ],
                  ),
                ),
              ),
              
              const SizedBox(height: 30),
              
              Container(
                padding: const EdgeInsets.all(16),
                decoration: BoxDecoration(
                  color: Colors.blue[50],
                  borderRadius: BorderRadius.circular(10),
                  border: Border.all(color: Colors.blue.shade200),
                ),
                child: const Row(
                  children: [
                    Icon(Icons.info_outline, color: Colors.blue),
                    SizedBox(width: 12),
                    Expanded(
                      child: Text(
                        'The Excel file will be created with headers. You can start adding student data after creation.',
                        style: TextStyle(
                          fontSize: 13,
                          color: Colors.black87,
                        ),
                      ),
                    ),
                  ],
                ),
              ),
              
              const SizedBox(height: 40),
              
              ElevatedButton.icon(
                onPressed: _isCreating ? null : _createExcelAndNavigate,
                icon: _isCreating
                    ? const SizedBox(
                        height: 20,
                        width: 20,
                        child: CircularProgressIndicator(
                          strokeWidth: 2,
                          valueColor: AlwaysStoppedAnimation<Color>(Colors.white),
                        ),
                      )
                    : const Icon(Icons.check_circle, size: 24),
                label: Text(
                  _isCreating ? 'Creating...' : 'Create Excel File',
                  style: const TextStyle(
                    fontSize: 18,
                    fontWeight: FontWeight.bold,
                  ),
                ),
                style: ElevatedButton.styleFrom(
                  padding: const EdgeInsets.symmetric(vertical: 16),
                  shape: RoundedRectangleBorder(
                    borderRadius: BorderRadius.circular(30),
                  ),
                  backgroundColor: Colors.blue,
                  foregroundColor: Colors.white,
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}
