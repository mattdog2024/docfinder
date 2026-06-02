# -*- coding: utf-8 -*-
"""
v1.8 核心逻辑测试脚本
测试：IndexEngine、IndexBuilder、搜索功能
"""
import sys
import os
import tempfile
import shutil
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def test_imports():
    """测试所有模块导入"""
    print("=== 测试模块导入 ===")
    try:
        from core.extractor import extract_text, SUPPORTED_EXTENSIONS
        print(f"  ✅ extractor 导入成功")
        print(f"  支持格式: {list(SUPPORTED_EXTENSIONS.keys())}")
    except Exception as e:
        print(f"  ❌ extractor 导入失败: {e}")
        traceback.print_exc()
        return False

    try:
        from core.indexer import IndexEngine, IndexBuilder
        print(f"  ✅ indexer 导入成功")
    except Exception as e:
        print(f"  ❌ indexer 导入失败: {e}")
        traceback.print_exc()
        return False

    try:
        import multiprocessing
        cpu_count = multiprocessing.cpu_count()
        print(f"  ✅ multiprocessing 导入成功，CPU 核心数: {cpu_count}")
    except Exception as e:
        print(f"  ❌ multiprocessing 导入失败: {e}")
        return False

    return True


def test_index_engine():
    """测试 IndexEngine 基本功能"""
    print("\n=== 测试 IndexEngine ===")
    from core.indexer import IndexEngine

    with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as f:
        db_path = f.name

    try:
        engine = IndexEngine(db_path)
        print(f"  ✅ IndexEngine 创建成功")

        info = engine.get_index_info()
        print(f"  ✅ get_index_info: {info}")

        # 测试写入和搜索
        records = [
            ('/path/to/test.docx', '测试文档', '这是一个测试文档，包含关键词：Python 搜索引擎', 
             'test.docx', 1024, 'abc123', '/path/to')
        ]
        engine.batch_add_documents(records, '/path/to')
        print(f"  ✅ batch_add_documents 成功")

        results = engine.search('Python')
        print(f"  ✅ search('Python') 返回 {len(results)} 条结果")
        if results:
            print(f"     第一条: {results[0]['filename']}")

        results = engine.search('搜索引擎')
        print(f"  ✅ search('搜索引擎') 返回 {len(results)} 条结果")

        return True
    except Exception as e:
        print(f"  ❌ IndexEngine 测试失败: {e}")
        traceback.print_exc()
        return False
    finally:
        os.unlink(db_path)


def test_index_builder():
    """测试 IndexBuilder.build_index 参数签名"""
    print("\n=== 测试 IndexBuilder ===")
    from core.indexer import IndexEngine, IndexBuilder
    import inspect

    sig = inspect.signature(IndexBuilder.build_index)
    params = list(sig.parameters.keys())
    print(f"  build_index 参数: {params}")

    # 验证参数名
    expected_params = ['self', 'root_dir', 'enabled_extensions', 'enable_pdf', 
                       'enable_ocr', 'max_workers', 'progress_callback', 
                       'log_callback', 'speed_callback']
    
    for p in expected_params:
        if p in params:
            print(f"  ✅ 参数 '{p}' 存在")
        else:
            print(f"  ❌ 参数 '{p}' 不存在！")

    # 测试实际调用
    with tempfile.NamedTemporaryFile(suffix='.db', delete=False) as f:
        db_path = f.name

    test_dir = tempfile.mkdtemp()
    try:
        # 创建一个测试文件（非 docx，不会被解析，但会被扫描）
        with open(os.path.join(test_dir, 'readme.txt'), 'w', encoding='utf-8') as f:
            f.write('测试文件')

        engine = IndexEngine(db_path)
        builder = IndexBuilder(engine)

        logs = []
        def log_cb(msg):
            logs.append(msg)

        stats = builder.build_index(
            root_dir=test_dir,
            enabled_extensions=['.docx', '.xlsx', '.pdf'],
            enable_pdf=False,
            enable_ocr=False,
            max_workers=2,
            progress_callback=None,
            log_callback=log_cb,
            speed_callback=None,
        )
        print(f"  ✅ build_index 调用成功，stats: {stats}")
        print(f"  日志: {logs}")
        return True
    except Exception as e:
        print(f"  ❌ build_index 调用失败: {e}")
        traceback.print_exc()
        return False
    finally:
        os.unlink(db_path)
        shutil.rmtree(test_dir)


def test_main_window_import():
    """测试 main_window 导入"""
    print("\n=== 测试 main_window 导入 ===")
    try:
        from ui.main_window import MainWindow, APP_VERSION, APP_NAME
        print(f"  ✅ main_window 导入成功")
        print(f"  版本号: {APP_VERSION}")
        print(f"  应用名: {APP_NAME}")
        return True
    except Exception as e:
        print(f"  ❌ main_window 导入失败: {e}")
        traceback.print_exc()
        return False


if __name__ == '__main__':
    print("DocFinder v1.8 核心逻辑测试")
    print("=" * 50)
    
    results = []
    results.append(("模块导入", test_imports()))
    results.append(("IndexEngine", test_index_engine()))
    results.append(("IndexBuilder", test_index_builder()))
    results.append(("main_window 导入", test_main_window_import()))
    
    print("\n" + "=" * 50)
    print("测试结果汇总:")
    all_passed = True
    for name, passed in results:
        status = "✅ 通过" if passed else "❌ 失败"
        print(f"  {status} - {name}")
        if not passed:
            all_passed = False
    
    print()
    if all_passed:
        print("🎉 所有测试通过！v1.8 核心逻辑正常")
    else:
        print("⚠️ 部分测试失败，请检查上方错误信息")
    
    sys.exit(0 if all_passed else 1)
