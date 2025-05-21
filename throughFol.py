import os
import shutil

def copy_files(src_dir, dest_dir):
    if not os.path.exists(dest_dir):
        os.makedirs(dest_dir)
    
    for root, _, files in os.walk(src_dir):
        for file in files:
            src_path = os.path.join(root, file)
            dest_path = os.path.join(dest_dir, file)
            
            # Tránh ghi đè nếu file trùng tên
            if os.path.exists(dest_path):
                base, ext = os.path.splitext(file)
                counter = 1
                while os.path.exists(dest_path):
                    new_name = f"{base}_{counter}{ext}"
                    dest_path = os.path.join(dest_dir, new_name)
                    counter += 1
            
            shutil.copy2(src_path, dest_path)
            print(f"Đã sao chép: {src_path} -> {dest_path}")

if __name__ == "__main__":
    source_directory = input("cvFile")
    destination_directory = input("result")
    
    copy_files(source_directory, destination_directory)
    print("Hoàn thành sao chép tất cả các file!")
