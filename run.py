#!/usr/bin/env python
"""
Excel Auto Launcher
Quick launcher script with helpful options
"""
import sys
import os
import argparse

def main():
    parser = argparse.ArgumentParser(
        description='Excel Auto - 엑셀 자동화 워크플로우 GUI 애플리케이션',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
예제:
  %(prog)s                    # GUI 실행
  %(prog)s --test             # 테스트 실행
  %(prog)s --example          # 사용 예제 실행
  %(prog)s --help             # 도움말 표시
        '''
    )
    
    parser.add_argument('--test', action='store_true',
                       help='테스트 스위트 실행')
    parser.add_argument('--example', action='store_true',
                       help='사용 예제 실행')
    parser.add_argument('--no-gui', action='store_true',
                       help='GUI 없이 실행 (테스트/예제용)')
    
    args = parser.parse_args()
    
    if args.test:
        print("테스트 실행 중...")
        import test_excel_auto
        return test_excel_auto.main()
    
    elif args.example:
        print("사용 예제 실행 중...")
        import usage_example
        usage_example.main()
        return 0
    
    else:
        # Run GUI
        print("Excel Auto GUI 시작 중...")
        try:
            import main
            main.main()
            return 0
        except ImportError as e:
            print(f"오류: {e}")
            print("\nGUI를 실행하려면 tkinter가 필요합니다.")
            print("Ubuntu/Debian: sudo apt-get install python3-tk")
            print("또는 --test 또는 --example 옵션을 사용하세요.")
            return 1
        except Exception as e:
            print(f"오류: {e}")
            import traceback
            traceback.print_exc()
            return 1

if __name__ == '__main__':
    sys.exit(main())
