import os
import time
import bypy
import pymysql


class BaiduPanWrapper:
    def __init__(self, api_key, secret_key, access_token, refresh_token):
        self.client = bypy.ByPy()
        self.client.config.api_key = api_key
        self.client.config.secret_key = secret_key
        self.client.config.access_token = access_token
        self.client.config.refresh_token = refresh_token

    def get_folder_files(self, folder_name):
        """
        获取指定文件夹内的所有文件，并只返回格式为.docx和.xlsx的文件
        """
        files = self.client.list_files(folder_name, recursive=True)['list']
        result = []
        for file in files:
            if not file['isdir'] and file['server_filename'].endswith('.docx') or file['server_filename'].endswith('.xlsx'):
                result.append(file)
        return result

    def download_file(self, file_name):
        """
        下载指定的文件
        """
        self.client.download(file_name)

    def delete_file(self, file_name):
        """
        删除指定的文件
        """
        self.client.delete(file_name)


class WordPressWrapper:
    def __init__(self, host, username, password, database):
        self.db = pymysql.connect(
            host=host,
            user=username,
            passwd=password,
            database=database
        )

    def publish_article(self, title, content):
        """
        发布一篇文章到WordPress
        """
        try:
            with self.db.cursor() as cursor:
                # 检查是否已经有相同标题的文章，防止重复发布
                sql = "SELECT * FROM wp_posts WHERE post_title = %s"
                cursor.execute(sql, (title,))
                articles = cursor.fetchall()
                if articles:
                    print(f'文章标题为 {title} 的文章已经存在，不再发布')
                    return
                # 插入文章数据到数据库
                sql = "INSERT INTO wp_posts (post_author, post_date, post_date_gmt, post_content, post_title, post_excerpt, post_status, comment_status, ping_status, post_password, post_name, to_ping, pinged, post_modified, post_modified_gmt, post_content_filtered, post_parent, guid, menu_order, post_type, post_mime_type, comment_count) VALUES (%s, NOW(), NOW(), %s, %s, '', %s, 'closed', 'closed', '', %s, '', '', NOW(), NOW(), '', 0, %s, 0, 'post', '', 0)"
                cursor.execute(sql, (1, content, title, 'publish', title, 'https://www.example.com/' + title + '/', title))
            self.db.commit()
            print(f'文章已发布，标题：{title}')
        except Exception as e:
            self.db.rollback()
            print('文章发布失败：', e)

    def __del__(self):
        self.db.close()


def get_file_content(file_name):
    """
    获取指定文件的内容
    """
    with open(file_name, 'rb') as f:
        content = f.read()
    return content


def extract_title(content):
    """
    根据文本内容提取文章标题，暂时使用文件名作为标题
    """
    file_name = os.path.basename(content.name)
    title = os.path.splitext(file_name)[0]
    return title


def check_new_file(api_key, secret_key, access_token, refresh_token, watch_folder, wp_host, wp_username, wp_password, wp_database):
    # 创建百度网盘和WordPress对象
    baidu_wrapper = BaiduPanWrapper(api_key, secret_key, access_token, refresh_token)
    wp_wrapper = WordPressWrapper(wp_host, wp_username, wp_password, wp_database)

    # 获取指定文件夹内的所有文件
    try:
        files = baidu_wrapper.get_folder_files(watch_folder)
    except bypy.BypyException as e:
        print('获取文件列表失败：', e)
        return

    # 遍历文件列表
    flag = False
    for file in files:
        try:
            # 判断文件的创建时间是否在10分钟内
            if (time.time() - file['server_ctime']) <= 600:
                flag = True
                # 下载文件
                print(f'开始下载文件 {file["server_filename"]}')
                baidu_wrapper.download_file(file['server_filename'])
                # 获取文件内容和标题
                with open(file['server_filename'], 'rb') as f:
                    content = f.read()
                title = extract_title(f)
                # 发布文章到WordPress
                wp_wrapper.publish_article(title, content)
                # 删除百度网盘上的文件
                baidu_wrapper.delete_file(file['path'])
        except Exception as e:
            print('处理文件时出错：', e)
        finally:
            # 删除本地的文件
            os.remove(file['server_filename'])

    if not flag:
        print('没有发现新文件')


def check_database(wp_host, wp_username, wp_password, wp_database):
    """
    检查数据库中是否存在指定表
    """
    try:
        db = pymysql.connect(
            host=wp_host,
            user=wp_username,
            passwd=wp_password,
            database=wp_database
        )
        cursor = db.cursor()
        cursor.execute("SHOW TABLES LIKE 'wp_posts'")
        result = cursor.fetchone()
        if result:
            print('数据库检查完成，数据表存在')
        else:
            print('数据库检查完成，数据表不存在')
            # 创建数据表
            cursor.execute("CREATE TABLE wp_posts (ID BIGINT(20) UNSIGNED NOT NULL AUTO_INCREMENT, post_author BIGINT(20) UNSIGNED NOT NULL DEFAULT 0, post_date DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00', post_date_gmt DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00', post_content LONGTEXT NOT NULL, post_title TEXT NOT NULL, post_excerpt TEXT NOT NULL, post_status VARCHAR(20) NOT NULL DEFAULT 'publish', comment_status VARCHAR(20) NOT NULL DEFAULT 'closed', ping_status VARCHAR(20) NOT NULL DEFAULT 'closed', post_password VARCHAR(255) NOT NULL DEFAULT '', post_name VARCHAR(200) NOT NULL DEFAULT '', to_ping TEXT NOT NULL, pinged TEXT NOT NULL, post_modified DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00', post_modified_gmt DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00', post_content_filtered LONGTEXT NOT NULL, post_parent BIGINT(20) UNSIGNED NOT NULL DEFAULT 0, guid VARCHAR(255) NOT NULL DEFAULT '', menu_order INT(11) NOT NULL DEFAULT 0, post_type VARCHAR(20) NOT NULL DEFAULT 'post', post_mime_type VARCHAR(100) NOT NULL DEFAULT '', comment_count BIGINT(20) NOT NULL DEFAULT 0, PRIMARY KEY (ID))")
            db.commit()
            print('数据表创建完成')
    except Exception as e:
        print('数据库检查失败：', e)
    finally:
        db.close()


if __name__ == '__main__':
    # 输入相关信息并检查数据库
    api_key = input('请输入AppKey：')
    secret_key = input('请输入SecretKey：')
    access_token = input('请输入Access Token：')
    refresh_token = input('请输入Refresh Token：')
    watch_folder = input('请输入监控的文件夹名：')
    wp_host = input('请输入WordPress的Host：')
    wp_username = input('请输入WordPress的用户名：')
    wp_password = input('请输入WordPress的密码：')
    wp_database = input('请输入WordPress的数据库名：')
    check_database(wp_host, wp_username, wp_password, wp_database)

    # 每隔10分钟监测一次新文件
    while True:
        check_new_file(api_key, secret_key, access_token, refresh_token, watch_folder, wp_host, wp_username, wp_password, wp_database)
        time.sleep(600)  # 休眠10分钟
