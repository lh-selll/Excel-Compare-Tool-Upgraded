import subprocess
import re
import hashlib
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad
import base64
import os
import sys
from PySide6.QtWidgets import QApplication, QMessageBox

class DeviceIDLicenseVerify():
    def __init__(self, license_file_path):
        self.license_file_path = license_file_path
        self.UUID = None

    def get_device_id(self):
        """获取当前计算机的设备ID"""
        result = None
        try:
            # 尝试使用wmic命令获取设备ID
            result = subprocess.run(
                ["wmic", "csproduct", "get", "uuid"],
                capture_output=True,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            self.UUID = result
            # 调试信息
            print(f"wmic命令输出: {result.stdout}")

            # 解析输出结果，跳过空行
            lines = [line.strip() for line in result.stdout.split('\n') if line.strip()]
            self.UUID = lines
            # 调试信息
            print(f"解析后的行: {lines}")

            if len(lines) >= 2:
                # 取第二行作为设备ID
                device_id = lines[1]

                # 调试信息
                print(f"提取的设备ID: {device_id}")
                self.UUID = device_id

                # 清理可能的额外字符
                device_id = re.sub(r'[^\w-]', '', device_id)
                self.UUID = device_id

                # 转换为标准UUID格式
                if re.match(r'^[0-9A-Fa-f]{32}$', device_id):
                    device_id = f"{device_id[:8]}-{device_id[8:12]}-{device_id[12:16]}-{device_id[16:20]}-{device_id[20:]}"
                    self.UUID = device_id

                # 验证设备ID格式
                if re.match(r'^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}$', device_id):
                    self.UUID = device_id
                    return device_id
                else:
                    error_msg = f"错误: 获取的设备ID格式不正确 - {device_id}"
                    self.show_error_message(error_msg)
                    return None
            else:
                error_msg = f"错误: 无法获取设备ID - lines = {lines}"
                self.show_error_message(error_msg)
                return None

        except Exception as e:
            error_msg = f"错误: 获取设备ID失败 - {str(e)}"
            self.show_error_message(error_msg)
            return None

    def verify_license(self):
        # 获取当前设备ID
        device_id = self.get_device_id()
        if not device_id:
            return False

        print(f"检测到设备ID: {device_id}")

        # 检查授权文件是否存在
        if not os.path.exists(self.license_file_path):
            error_msg = f"错误: 未找到授权文件 {self.license_file_path}"
            self.show_error_message(error_msg)
            return False

        # 读取授权文件
        try:
            with open(self.license_file_path, 'r') as f:
                license_data = f.read().strip().split(',')
                if len(license_data) != 2:
                    error_msg = "错误: 授权文件格式不正确"
                    self.show_error_message(error_msg)
                    return False

                iv = base64.b64decode(license_data[0])
                ciphertext = base64.b64decode(license_data[1])
        except Exception as e:
            error_msg = f"错误: 读取授权文件时发生异常: {e}"
            self.show_error_message(error_msg)
            return False

        # 使用SHA256对设备ID进行哈希
        try:
            device_id_hash = hashlib.sha256(device_id.encode()).digest()
        except Exception as e:
            error_msg = f"错误: 对设备ID进行哈希时发生异常: {e}"
            self.show_error_message(error_msg)
            return False

        # 截取前16字节作为AES密钥(128位)
        aes_key = device_id_hash[:16]

        # 尝试解密
        try:
            cipher = AES.new(aes_key, AES.MODE_CBC, iv)
            decrypted_data = unpad(cipher.decrypt(ciphertext), AES.block_size)

            # 验证解密结果
            if decrypted_data.decode() == "授权成功":
                print("授权验证成功!")
                return True
            else:
                error_msg = "错误: 授权验证失败"
                self.show_error_message(error_msg)
                return False
        except Exception as e:
            error_msg = f"错误: 授权验证过程中发生异常: {e}"
            self.show_error_message(error_msg)
            return False

    def show_error_message(self, message):
        # root = tk.Tk()
        # root.withdraw()  # 隐藏主窗口
        # messagebox.showerror("授权验证失败", f"{message}\n授权验证失败，程序无法继续运行。\n读取的UUID为: {self.UUID}\n请联系开发者获取有效的授权文件。")# 创建临时 QApplication 实例（如果不存在）
        if not QApplication.instance():
            app = QApplication(sys.argv)
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("授权验证失败")
        msg.setInformativeText(f"{message}\n授权验证失败，程序无法继续运行。\n读取的UUID为: {self.UUID}\n请联系开发者获取有效的授权文件。")
        msg.setWindowTitle("错误信息")
        msg.exec()

def main():
    print("Excel对比工具 - 授权验证")
    print("=" * 40)
    license_path = ".\\license\\license.key"
    Verify_app = DeviceIDLicenseVerify(license_path)

    # 验证授权
    if not Verify_app.verify_license():
        print("=" * 40)
        print("授权验证失败，程序无法继续运行。")
        print("请联系管理员获取有效的授权文件。")
        sys.exit(1)

    # 授权通过，运行主程序
    print("=" * 40)
    print("Hello World!")
    print("这是一个经过授权的程序。")
    print("=" * 40)

    # 这里可以添加你的其他功能代码

if __name__ == "__main__":
    main()