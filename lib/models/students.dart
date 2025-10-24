class Student {
  final String name;
  final String phoneNumber;

  Student({
    required this.name,
    required this.phoneNumber,
  });

  // Convert Student object to Map for Excel
  Map<String, dynamic> toMap() {
    return {
      'Name': name,
      'Phone Number': phoneNumber,
    };
  }

  // Create Student object from Map
  factory Student.fromMap(Map<String, dynamic> map) {
    return Student(
      name: map['Name'] ?? '',
      phoneNumber: map['Phone Number'] ?? '',
    );
  }

  @override
  String toString() {
    return 'Student{name: $name, phoneNumber: $phoneNumber}';
  }
}