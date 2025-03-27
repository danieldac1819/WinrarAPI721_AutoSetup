import os
import subprocess
import requests
import winreg as reg
import re
import sys
import time
import platform
import webbrowser 

# Hàm lấy đường dẫn cài đặt WinRAR từ Registry

def get_winrar_path():
    try:
        # Kiểm tra xem hệ điều hành đang sử dụng 64-bit hay 32-bit
        is_64bit = platform.architecture()[0] == '64bit'

        # Định nghĩa các registry key cho cả phiên bản 32-bit và 64-bit
        if is_64bit:
            reg_paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe",  # 64-bit registry path
                r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe"  # 32-bit registry path
            ]
        else:
            reg_paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe"  # 32-bit registry path
            ]
        
        # Lặp qua các registry key để tìm WinRAR
        for path in reg_paths:
            try:
                reg_key = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, path)
                winrar_path, _ = reg.QueryValueEx(reg_key, "")
                reg.CloseKey(reg_key)
                return winrar_path  # Trả về đường dẫn WinRAR nếu tìm thấy
            except OSError:
                continue  # Nếu không tìm thấy, tiếp tục kiểm tra các khoá khác
        
        return None  # Nếu không tìm thấy WinRAR trong cả hai khoá, trả về None
    except Exception as e:
        print(f"Error: {e}")
        return None


# Hàm kiểm tra phiên bản WinRAR hiện tại từ hệ thống
# Hàm kiểm tra phiên bản WinRAR hiện tại từ hệ thống
def check_rar_version(winrar_path):
    rar_path = os.path.join(winrar_path, "rar.exe")
    if not os.path.exists(rar_path):
        return f"Không tìm thấy rar.exe trong đường dẫn: {winrar_path}"

    try:
        result = subprocess.run([rar_path, "-iver"], capture_output=True, text=True)
        if result.returncode == 0:
            return result.stdout.strip()  # Trả về thông tin phiên bản đầy đủ
        return "Không thể lấy phiên bản từ rar.exe."
    except Exception as e:
        return f"Không thể chạy rar.exe: {str(e)}"


# Hàm lấy thông tin phiên bản và liên kết từ API
def get_version_and_link_from_url():
    url = "https://api.itdev721.workers.dev/?action=WinrarVersionJson"
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Raise exception for HTTP errors
        data = response.json()

        # Trả về thông tin tải xuống WinRAR cho 64-bit
        number_version_current = data.get('NumberversionCurrent')
        version_current = data.get('VersionCurrent')
        link_current = data.get('LinkCurrent')

        # Trả về thông tin tải xuống WinRAR cho 32-bit
        version_current32 = data.get('Versioncurrent32')
        link_current32 = data.get('LinkCurrent32')

        return number_version_current, version_current, link_current, version_current32, link_current32
    except (requests.exceptions.RequestException, ValueError) as e:
        print(f"Đã xảy ra lỗi khi tải dữ liệu từ API: {str(e)}")
        return None, None, None, None, None


# Hàm kiểm tra hệ điều hành 64-bit
def is_64bit_os():
    # Lấy đường dẫn đến ổ đĩa hệ thống (thường là C:\ trên Windows)
    c_drive = os.environ['SystemDrive']
    
    # Đường dẫn đến thư mục Program Files (x86)
    program_files_86 = os.path.join(c_drive, "Program Files (x86)")
    
    # Kiểm tra nếu thư mục Program Files (x86) tồn tại
    if os.path.exists(program_files_86):
        return True  # Hệ điều hành 64-bit
    else:
        return False  # Hệ điều hành 32-bit


# Hàm làm sạch phiên bản (loại bỏ các hậu tố như x64, x86)
def clean_local_version(local_version):
    return re.sub(r'\s*(x86|x64|32|bits)?\s*$', '', local_version).strip()


# Hàm tải xuống tệp từ URL hiển thị tiến độ
def download_file(download_url, save_path, retries=3):
    attempt = 0
    while attempt < retries:
        try:
            response = requests.get(download_url, stream=True)
            response.raise_for_status()
            
            # Lấy kích thước của tệp (nếu có)
            total_size = int(response.headers.get('content-length', 0))
            with open(save_path, 'wb') as file:
                downloaded = 0
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        file.write(chunk)
                        downloaded += len(chunk)
                        
                        # Hiển thị tiến độ tải xuống
                        if total_size > 0:
                            progress = int(50 * downloaded / total_size)  # Tính toán tiến độ (mức độ 50)
                            bar = '█' * progress + '-' * (50 - progress)  # Biểu diễn thanh tiến độ
                            print(f"\r[{bar}] {downloaded / 1024:.2f} KB / {total_size / 1024:.2f} KB", end='')

            print(f"\nTải xuống hoàn tất: {save_path}")
            return
        except requests.exceptions.RequestException as e:
            attempt += 1
            print(f"\nĐã xảy ra lỗi khi tải xuống (thử lại {attempt}/{retries}): {str(e)}")
            if attempt == retries:
                print("Không thể tải xuống sau nhiều lần thử.")


# Hàm chạy trình cài đặt WinRAR mới trong chế độ cài đặt thầm lặng (silent)
def run_installer(installer_path):
    try:
        command = [installer_path, "/S", ""]  # Chế độ cài đặt thầm lặng
        subprocess.run(command, check=True, shell=True)
    except subprocess.CalledProcessError as e:
        print(f"Đã xảy ra lỗi khi chạy trình cài đặt: {str(e)}")
    except Exception as e:
        print(f"Đã xảy ra lỗi không xác định: {str(e)}")


# Hàm kiểm tra và tải tệp KeyCopyRight vào thư mục cài đặt WinRAR
def download_key_copy_right(winrar_folder, key_copy_right_url):
    key_copy_right_filename = os.path.basename(key_copy_right_url)
    key_copy_right_path = os.path.join(winrar_folder, key_copy_right_filename)

    if not os.path.exists(key_copy_right_path):
        print(f"Tệp KeyCopyRight không tồn tại. Đang tải xuống từ: {key_copy_right_url}")
        download_file(key_copy_right_url, key_copy_right_path)
    else:
        print("Tệp KeyCopyRight đã tồn tại.")


# Hàm lấy thư mục cài đặt WinRAR (bao gồm cả tên tập tin WinRAR.exe)
def get_folder_setup_winrar():
    winrar_path = get_winrar_path()
    if winrar_path:
        # Trả về thư mục chứa WinRAR.exe
        return os.path.dirname(winrar_path)  # Lấy thư mục từ đường dẫn WinRAR.exe
    else:
        return None


# Hàm hiển thị trợ giúp và tạo tệp với tên {script_name}.cmd nếu không có
def print_help():
    script_name = os.path.basename(sys.argv[0])  # Lấy tên tệp hiện tại (ví dụ: "WinrarAPI721_AutoSetup.py")
    cmd_file_name = f"{os.path.splitext(script_name)[0]}.cmd"  # Tạo tên tệp .cmd (ví dụ: "WinrarAPI721_AutoSetup.cmd")

    # Lấy thông tin cài đặt WinRAR và phiên bản
    winrar_folder = get_folder_setup_winrar()  # Sử dụng get_folder_setup_winrar() để lấy đường dẫn thư mục cài đặt
    winrar_version = ""
    
    if winrar_folder:
        winrar_version = check_rar_version(winrar_folder)  # Kiểm tra phiên bản WinRAR từ thư mục cài đặt
    else:
        winrar_version = "Không tìm thấy WinRAR cài đặt trên hệ thống."

    help_text = f"""
    *******************721PC-Net Corporation***********************
    
    ** Phiên bản nâng cấp sửa lỗi tập tin cmd không hổ trợ utf-8
    
    ** Trợ giúp: Các tham số dòng lệnh có sẵn cho chương trình **

    ***************************************************************

    ** Các tham số dòng lệnh bổ sung: **
    1. /autosetup    : Tự động cài đặt WinRAR trên hệ thống.
    2. /autosetup 32 : Cài đặt WinRAR phiên bản 32-bit.
    3. /autosetup 64 : Cài đặt WinRAR phiên bản 64-bit.
    4. /autosetup /help : truy cập vào facepage hổ trợ
    ** Cách sử dụng cho /autosetup: **
    - Chạy script với tham số `/autosetup` để tự động cài đặt WinRAR trên hệ thống của bạn.
    - Ví dụ:
        - {script_name} /autosetup
        - {script_name} /autosetup 32 (dành cho hệ điều hành 32-bit)
        - {script_name} /autosetup 64 (dành cho hệ điều hành 64-bit)

    ***************************************************************
    
    ** Cách sử dụng: **
    mở tập tin:  {cmd_file_name}

    hoặc tích hợp vào ứng dụng của bạn với tham số:

    {script_name} /autosetup

    phát triển API bởi wWw.721PC.Asia <<>> 721PC-Net Corporation

    truy cập facepage: fb.com/721PC

    ***************************************************************
    
    ** Thông tin cài đặt WinRAR: **
    - Thư mục cài đặt WinRAR: {winrar_folder if winrar_folder else 'Không tìm thấy thư mục cài đặt.'}
    - Phiên bản WinRAR hiện tại: {winrar_version}
    
    ***************************************************************
    """

    # Kiểm tra thư mục root có tệp {script_name}.cmd không
    root_dir = os.getcwd()  # Lấy thư mục hiện tại
    cmd_file_path = os.path.join(root_dir, cmd_file_name)  # Đặt tên tệp là {script_name}.cmd

    # Nếu tệp .cmd không tồn tại, tạo nó với mã hóa ANSI (CP1252)
    if not os.path.exists(cmd_file_path):
        print(f"Không tìm thấy tệp \n{cmd_file_path}.\n Đang tạo tệp CMD để tự động cài đặt WinRAR...")
        # Mở tệp với mã hóa ANSI (CP1252)
        with open(cmd_file_path, "w", encoding="cp1252") as f:
            # Ghi nội dung vào tệp .cmd
            f.write(f"{script_name} /autosetup\n\n\n")
            f.write(f":: Cài đặt auto setup với thiết lập mặc định\n")
            f.write(f":: {script_name} /autosetup\n\n")

            f.write(f":: Cài đặt auto setup với phiên bản 32-bit\n")
            f.write(f":: {script_name} /autosetup 32\n\n")

            f.write(f":: Cài đặt auto setup với phiên bản 64-bit\n")
            f.write(f":: {script_name} /autosetup 64\n")
        
        print(f"Tạo tệp CMD: {cmd_file_path} thành công!")

    # In ra nội dung trợ giúp
    print(help_text)


def main():
    # Kiểm tra tham số dòng lệnh
    if '/autosetup' in sys.argv:
        if len(sys.argv) == 3 and sys.argv[2] == "64":
            print("Đang tải phiên bản 64-bit WinRAR...")
            # Lấy thông tin phiên bản và liên kết tải xuống cho 64-bit
            number_version_current, version_current, link_current, version_current32, link_current32 = get_version_and_link_from_url()
            
            if number_version_current and version_current and link_current:
                print(f"Tải xuống và cài đặt WinRAR 64-bit phiên bản mới: {version_current}")
                
                # Tải xuống trình cài đặt WinRAR 64-bit
                save_path = f"{version_current}.exe"
                download_file(link_current, save_path)
                
                # Chạy trình cài đặt
                run_installer(save_path)
                
                # Xóa tập tin cài đặt sau khi cài đặt thành công
                if os.path.exists(save_path):
                    os.remove(save_path)
                    print(f"Tập tin {save_path} đã được xóa.")
            else:
                print("Không thể lấy thông tin phiên bản hoặc liên kết tải xuống từ API.")
        
        elif len(sys.argv) == 3 and sys.argv[2] == "32":
            print("Đang tải phiên bản 32-bit WinRAR...")
            # Lấy thông tin phiên bản và liên kết tải xuống cho 32-bit
            number_version_current, version_current, link_current, version_current32, link_current32 = get_version_and_link_from_url()
            
            if version_current32 and link_current32:
                print(f"Tải xuống và cài đặt WinRAR 32-bit phiên bản mới: {version_current32}")
                
                # Tải xuống trình cài đặt WinRAR 32-bit
                save_path = f"{version_current32}.exe"
                download_file(link_current32, save_path)
                
                # Chạy trình cài đặt
                run_installer(save_path)
                
                # Xóa tập tin cài đặt sau khi cài đặt thành công
                if os.path.exists(save_path):
                    os.remove(save_path)
                    print(f"Tập tin {save_path} đã được xóa.")
            else:
                print("Không thể lấy thông tin phiên bản hoặc liên kết tải xuống từ API.")
        
        else:
            # Kiểm tra hệ điều hành hiện tại và tải phiên bản phù hợp (32-bit hay 64-bit)
            if is_64bit_os():
                print("Hệ điều hành 64-bit, tải xuống phiên bản 64-bit WinRAR...")
                number_version_current, version_current, link_current, version_current32, link_current32 = get_version_and_link_from_url()
                
                if number_version_current and version_current and link_current:
                    print(f"Tải xuống và cài đặt WinRAR 64-bit phiên bản mới: {version_current}")
                    
                    # Tải xuống trình cài đặt WinRAR 64-bit
                    save_path = f"{version_current}.exe"
                    download_file(link_current, save_path)
                    
                    # Chạy trình cài đặt
                    run_installer(save_path)
                    
                    # Xóa tập tin cài đặt sau khi cài đặt thành công
                    if os.path.exists(save_path):
                        os.remove(save_path)
                        print(f"Tập tin {save_path} đã được xóa.")
                else:
                    print("Không thể lấy thông tin phiên bản hoặc liên kết tải xuống từ API.")
            else:
                print("Hệ điều hành 32-bit, tải xuống phiên bản 32-bit WinRAR...")
                number_version_current, version_current, link_current, version_current32, link_current32 = get_version_and_link_from_url()
                
                if version_current32 and link_current32:
                    print(f"Tải xuống và cài đặt WinRAR 32-bit phiên bản mới: {version_current32}")
                    
                    # Tải xuống trình cài đặt WinRAR 32-bit
                    save_path = f"{version_current32}.exe"
                    download_file(link_current32, save_path)
                    
                    # Chạy trình cài đặt
                    run_installer(save_path)
                    
                    # Xóa tập tin cài đặt sau khi cài đặt thành công
                    if os.path.exists(save_path):
                        os.remove(save_path)
                        print(f"Tập tin {save_path} đã được xóa.")
                else:
                    print("Không thể lấy thông tin phiên bản hoặc liên kết tải xuống từ API.")
    
    elif '/help' in sys.argv:
        print_help()
        print("Mở trang Facepage fb.com/721PC...")
        webbrowser.open('https://fb.com/721PC')
        sys.exit(0)  # Thoát sau khi hiển thị trợ giúp
    
    else:
        print_help()
        input("Nhấn Enter để thoát...")  # Chờ người dùng nhấn Enter để thoát


if __name__ == "__main__":
    main()
