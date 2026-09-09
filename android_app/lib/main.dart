import 'package:flutter/material.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:webview_flutter/webview_flutter.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
  runApp(const PhanCongApp());
}

class PhanCongApp extends StatelessWidget {
  const PhanCongApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'Phân công V10.3',
      theme: ThemeData(useMaterial3: true),
      home: const HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  const HomePage({super.key});

  @override
  State<HomePage> createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  static const defaultUrl = String.fromEnvironment('APP_URL', defaultValue: '');
  final _urlController = TextEditingController();
  WebViewController? _webController;
  bool _loadingSettings = true;
  String? _error;

  @override
  void initState() {
    super.initState();
    _loadSavedUrl();
  }

  Future<void> _loadSavedUrl() async {
    final prefs = await SharedPreferences.getInstance();
    final saved = prefs.getString('server_url') ?? defaultUrl;
    _urlController.text = saved;
    if (saved.isNotEmpty) {
      await _open(saved, save: false);
    }
    if (mounted) setState(() => _loadingSettings = false);
  }

  Future<void> _open(String raw, {bool save = true}) async {
    var url = raw.trim();
    if (url.isEmpty) {
      setState(() => _error = 'Hãy nhập địa chỉ web của phần mềm.');
      return;
    }
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      url = 'https://$url';
      _urlController.text = url;
    }
    final uri = Uri.tryParse(url);
    if (uri == null || uri.host.isEmpty) {
      setState(() => _error = 'Địa chỉ web không hợp lệ.');
      return;
    }

    if (save) {
      final prefs = await SharedPreferences.getInstance();
      await prefs.setString('server_url', url);
    }

    final controller = WebViewController()
      ..setJavaScriptMode(JavaScriptMode.unrestricted)
      ..setNavigationDelegate(
        NavigationDelegate(
          onPageStarted: (_) => mounted ? setState(() => _error = null) : null,
          onWebResourceError: (e) {
            if (mounted && e.isForMainFrame == true) {
              setState(() => _error = 'Không kết nối được máy chủ: ${e.description}');
            }
          },
        ),
      )
      ..loadRequest(uri);

    setState(() {
      _webController = controller;
      _error = null;
    });
  }

  void _showSettings() {
    setState(() => _webController = null);
  }

  @override
  Widget build(BuildContext context) {
    if (_loadingSettings) {
      return const Scaffold(body: Center(child: CircularProgressIndicator()));
    }

    final controller = _webController;
    if (controller != null) {
      return Scaffold(
        appBar: AppBar(
          title: const Text('Phân công V10.3'),
          actions: [
            IconButton(
              tooltip: 'Tải lại',
              onPressed: () => controller.reload(),
              icon: const Icon(Icons.refresh),
            ),
            IconButton(
              tooltip: 'Đổi máy chủ',
              onPressed: _showSettings,
              icon: const Icon(Icons.settings),
            ),
          ],
        ),
        body: Column(
          children: [
            if (_error != null)
              MaterialBanner(
                content: Text(_error!),
                actions: [
                  TextButton(onPressed: () => controller.reload(), child: const Text('THỬ LẠI')),
                  TextButton(onPressed: _showSettings, child: const Text('ĐỔI URL')),
                ],
              ),
            Expanded(child: WebViewWidget(controller: controller)),
          ],
        ),
      );
    }

    return Scaffold(
      appBar: AppBar(title: const Text('Phân công V10.3')),
      body: Center(
        child: ConstrainedBox(
          constraints: const BoxConstraints(maxWidth: 560),
          child: Padding(
            padding: const EdgeInsets.all(24),
            child: Column(
              mainAxisSize: MainAxisSize.min,
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                const Icon(Icons.assignment_ind_outlined, size: 72),
                const SizedBox(height: 18),
                const Text(
                  'Kết nối phần mềm Phân công chuyên môn',
                  textAlign: TextAlign.center,
                  style: TextStyle(fontSize: 22, fontWeight: FontWeight.w600),
                ),
                const SizedBox(height: 20),
                TextField(
                  controller: _urlController,
                  keyboardType: TextInputType.url,
                  autocorrect: false,
                  decoration: const InputDecoration(
                    border: OutlineInputBorder(),
                    labelText: 'Địa chỉ máy chủ',
                    hintText: 'https://ten-phan-mem.onrender.com',
                  ),
                  onSubmitted: (v) => _open(v),
                ),
                if (_error != null) ...[
                  const SizedBox(height: 10),
                  Text(_error!, style: TextStyle(color: Theme.of(context).colorScheme.error)),
                ],
                const SizedBox(height: 16),
                FilledButton.icon(
                  onPressed: () => _open(_urlController.text),
                  icon: const Icon(Icons.open_in_browser),
                  label: const Text('MỞ PHẦN MỀM'),
                ),
                const SizedBox(height: 10),
                const Text(
                  'Địa chỉ sẽ được lưu trên thiết bị. Có thể đổi lại bằng nút Cài đặt trong ứng dụng.',
                  textAlign: TextAlign.center,
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
