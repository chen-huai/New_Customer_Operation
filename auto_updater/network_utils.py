# -*- coding: utf-8 -*-
"""
网络连接检测工具
提供网络连接状态检查和诊断功能
"""

import socket
import requests
import time
from typing import Tuple, Optional, Dict, Any
from .config import (
    GITHUB_API_BASE,
    CONNECTION_TIMEOUT,
    CHECK_TIMEOUT,
    REQUEST_HEADERS
)

class NetworkConnectivityChecker:
    """网络连接检测器"""

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update(REQUEST_HEADERS)

    def check_dns_resolution(self, hostname: str = "api.github.com") -> Tuple[bool, str]:
        """
        检查DNS解析
        :param hostname: 要检测的主机名
        :return: (是否成功, 详细信息)
        """
        try:
            start_time = time.time()
            ip_address = socket.gethostbyname(hostname)
            resolve_time = round((time.time() - start_time) * 1000, 2)

            return True, f"DNS解析成功，{hostname} -> {ip_address} (耗时: {resolve_time}ms)"
        except socket.gaierror as e:
            return False, f"DNS解析失败: {str(e)}"
        except Exception as e:
            return False, f"DNS检测异常: {str(e)}"

    def check_tcp_connection(self, hostname: str = "api.github.com", port: int = 443) -> Tuple[bool, str]:
        """
        检查TCP连接
        :param hostname: 主机名
        :param port: 端口
        :return: (是否成功, 详细信息)
        """
        try:
            start_time = time.time()
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(CONNECTION_TIMEOUT)

            result = sock.connect_ex((hostname, port))
            connect_time = round((time.time() - start_time) * 1000, 2)

            sock.close()

            if result == 0:
                return True, f"TCP连接成功，{hostname}:{port} (耗时: {connect_time}ms)"
            else:
                return False, f"TCP连接失败，{hostname}:{port} (错误代码: {result})"

        except socket.timeout:
            return False, f"TCP连接超时（{CONNECTION_TIMEOUT}秒）: {hostname}:{port}"
        except Exception as e:
            return False, f"TCP连接异常: {str(e)}"

    def check_http_connection(self, url: str = None) -> Tuple[bool, str]:
        """
        检查HTTP连接
        :param url: 要检测的URL
        :return: (是否成功, 详细信息)
        """
        if url is None:
            url = f"{GITHUB_API_BASE}/"

        try:
            start_time = time.time()
            response = self.session.get(url, timeout=(CONNECTION_TIMEOUT, CHECK_TIMEOUT))
            response_time = round((time.time() - start_time) * 1000, 2)

            if response.status_code == 200:
                return True, f"HTTP连接成功，{url} (状态码: {response.status_code}, 耗时: {response_time}ms)"
            else:
                return False, f"HTTP响应异常，{url} (状态码: {response.status_code}, 耗时: {response_time}ms)"

        except requests.exceptions.Timeout:
            return False, f"HTTP连接超时（{CHECK_TIMEOUT}秒）: {url}"
        except requests.exceptions.ConnectionError as e:
            return False, f"HTTP连接失败: {str(e)}"
        except Exception as e:
            return False, f"HTTP检测异常: {str(e)}"

    def check_github_api_access(self) -> Tuple[bool, str]:
        """
        检查GitHub API访问权限
        :return: (是否成功, 详细信息)
        """
        try:
            url = f"{GITHUB_API_BASE}/rate_limit"
            start_time = time.time()
            response = self.session.get(url, timeout=(CONNECTION_TIMEOUT, CHECK_TIMEOUT))
            response_time = round((time.time() - start_time) * 1000, 2)

            if response.status_code == 200:
                data = response.json()
                resources = data.get('resources', {})
                core = resources.get('core', {})

                remaining = core.get('remaining', 0)
                limit = core.get('limit', 5000)
                used = limit - remaining

                return True, f"GitHub API访问正常 (已用: {used}/{limit}, 剩余: {remaining}, 耗时: {response_time}ms)"
            elif response.status_code == 403:
                return False, f"GitHub API访问受限 (状态码: {response.status_code})"
            else:
                return False, f"GitHub API访问失败 (状态码: {response.status_code})"

        except Exception as e:
            return False, f"GitHub API检测异常: {str(e)}"

    def measure_network_speed(self) -> Tuple[float, str]:
        """
        测量网络速度
        :return: (速度KB/s, 描述信息)
        """
        try:
            # 使用GitHub的一个小文件来测试速度
            test_url = "https://raw.githubusercontent.com/octocat/Hello-World/master/README"

            start_time = time.time()
            response = self.session.get(test_url, stream=True, timeout=10)
            response.raise_for_status()

            downloaded = 0
            test_size = 1024  # 只下载1KB用于测试

            for chunk in response.iter_content(chunk_size=128):
                if chunk:
                    downloaded += len(chunk)
                    if downloaded >= test_size:
                        break

            end_time = time.time()
            duration = end_time - start_time

            if duration > 0:
                speed_kb_per_sec = round((downloaded / 1024) / duration, 2)
                if speed_kb_per_sec > 100:
                    return speed_kb_per_sec, f"网络速度良好: {speed_kb_per_sec} KB/s"
                elif speed_kb_per_sec > 50:
                    return speed_kb_per_sec, f"网络速度一般: {speed_kb_per_sec} KB/s"
                else:
                    return speed_kb_per_sec, f"网络速度较慢: {speed_kb_per_sec} KB/s"
            else:
                return 0.0, "无法测量网络速度"

        except Exception as e:
            return 0.0, f"网络速度测试失败: {str(e)}"

    def comprehensive_network_check(self) -> Dict[str, Any]:
        """
        综合网络检测
        :return: 检测结果字典
        """
        results = {
            'overall_status': 'unknown',
            'checks': {},
            'recommendations': []
        }

        # 1. DNS检测
        dns_success, dns_info = self.check_dns_resolution()
        results['checks']['dns'] = {
            'success': dns_success,
            'info': dns_info
        }

        # 2. TCP连接检测
        tcp_success, tcp_info = self.check_tcp_connection()
        results['checks']['tcp'] = {
            'success': tcp_success,
            'info': tcp_info
        }

        # 3. HTTP连接检测
        http_success, http_info = self.check_http_connection()
        results['checks']['http'] = {
            'success': http_success,
            'info': http_info
        }

        # 4. GitHub API检测
        api_success, api_info = self.check_github_api_access()
        results['checks']['github_api'] = {
            'success': api_success,
            'info': api_info
        }

        # 5. 网络速度测试
        speed, speed_info = self.measure_network_speed()
        results['checks']['network_speed'] = {
            'success': speed > 0,
            'speed_kb_s': speed,
            'info': speed_info
        }

        # 判断整体状态
        success_count = sum(1 for check in results['checks'].values() if check['success'])
        total_count = len(results['checks'])

        if success_count == total_count:
            results['overall_status'] = 'excellent'
        elif success_count >= total_count * 0.8:
            results['overall_status'] = 'good'
        elif success_count >= total_count * 0.5:
            results['overall_status'] = 'poor'
        else:
            results['overall_status'] = 'failed'

        # 生成建议
        results['recommendations'] = self._generate_recommendations(results['checks'])

        return results

    def _generate_recommendations(self, checks: Dict) -> list:
        """
        根据检测结果生成建议
        :param checks: 检测结果
        :return: 建议列表
        """
        recommendations = []

        if not checks['dns']['success']:
            recommendations.append("请检查DNS设置，尝试更换DNS服务器（如8.8.8.8或1.1.1.1）")
            recommendations.append("检查网络连接是否正常，确认可以访问其他网站")

        if not checks['tcp']['success']:
            recommendations.append("检查防火墙设置，确保允许HTTPS连接（端口443）")
            recommendations.append("如果是公司网络，请联系网络管理员")

        if not checks['http']['success']:
            recommendations.append("检查代理服务器设置")
            recommendations.append("尝试使用其他网络环境")

        if not checks['github_api']['success']:
            recommendations.append("GitHub API可能暂时不可用，请稍后重试")
            recommendations.append("检查是否达到API请求频率限制")

        speed_check = checks['network_speed']
        if not speed_check['success'] or speed_check['speed_kb_s'] < 50:
            recommendations.append("网络速度较慢，建议在网络状况良好时进行更新")
            recommendations.append("可以考虑使用有线网络连接")

        if not recommendations:
            recommendations.append("网络连接正常，可以继续更新")

        return recommendations

def quick_connectivity_check() -> Tuple[bool, str]:
    """
    快速连接检测
    :return: (是否连接正常, 详细信息)
    """
    checker = NetworkConnectivityChecker()
    success, info = checker.check_dns_resolution()
    if not success:
        return False, f"网络连接问题: {info}"

    success, info = checker.check_github_api_access()
    if not success:
        return False, f"GitHub访问问题: {info}"

    return True, "网络连接正常，可以继续更新"