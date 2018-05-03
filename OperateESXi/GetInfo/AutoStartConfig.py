# -*- coding: utf-8 -*-
from pyVmomi import vim
from pyvim import connect

import time

from GetInfo.models import ClusterInfo


def auto_start(cluster_ip, start_opt):
    filter_hosts = ClusterInfo.objects.filter(IP=cluster_ip)

    if len(filter_hosts) == 1:
        t_host = ClusterInfo.objects.get(IP=cluster_ip)
        user = t_host.User
        pwd = t_host.Pwd
        mysi = connect.ConnectNoSSL(host=cluster_ip, port=443, user=user, pwd=pwd)
        container = mysi.content.viewManager.CreateContainerView(mysi.content.rootFolder, [vim.VirtualMachine], True)
        vm_obj = container.view[0]
        host_obj = vm_obj.summary.runtime.host
        hostdefsettings = vim.host.AutoStartManager.SystemDefaults()
        if start_opt is True:
            hostdefsettings.enabled = True
        else:
            hostdefsettings.enabled = False
        spec = host_obj.configManager.autoStartManager.config
        spec.defaults = hostdefsettings
        host_obj.configManager.autoStartManager.ReconfigureAutostart(spec)

        count = 0
        while count < 3:
            time.sleep(7)
            if host_obj.config.autoStart.defaults.enabled is start_opt:
                break
            count += 1

        auto_config = host_obj.config.autoStart.defaults.enabled
        connect.Disconnect(mysi)
        if auto_config is start_opt:
            return {'success': 'change vms auto start config success, now state is %s' % auto_config}
        else:
            return {'fail': 'change vms auto start config failed, cluster is %s - from auto_start' % cluster_ip}
    elif len(filter_hosts) > 1:
        return {'error': 'inner Error with: %s - from auto_start' % cluster_ip}
    else:
        return {'error': 'This IP: %s Not Found - from auto_start' % cluster_ip}
