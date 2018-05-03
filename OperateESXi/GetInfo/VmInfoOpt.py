# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from __future__ import print_function
from GetInfo.models import VmInfo, OldVmInfo, ClusterInfo
from GetInfo.VmStatusOpt import get_vms_status
import datadiff


def collect_vm_info(vm):
    # Get vm info from db
    # 根据cluster中vm的名称,从db中查询到d_myinfo,再获取db中的信息
    # not be used temporarily
    my_info = dict()
    my_info['IP'] = vm.Name
    my_info['IP'] = vm.Cluster
    my_info['Path'] = vm.vmPathName
    my_info['Guest'] = vm.guestFullName
    my_info['Annotation'] = vm.annotation
    my_info['instanceUuid'] = vm.instanceUuid
    my_info['PowerStatus'] = vm.powerStatus
    my_info['IP'] = vm.IP
    return my_info


def get_vms_info(cluster_ip):
    # show vms(all the vms in the cluster) info from db
    filter_hosts = ClusterInfo.objects.filter(IP=cluster_ip)

    if len(filter_hosts) == 1:
        t_vms = VmInfo.objects.filter(Cluster=cluster_ip)

        if len(t_vms) < 1:
            return {'error': 'No matched vms in %s - from get_vms_info' % cluster_ip}
        else:
            info_list = list(t_vms.values())
            return info_list
    else:
        return {'error': 'This IP: %s Not found in db - from get_vms_info' % cluster_ip}


def reset_vms_info(cluster_ip, reset=True):
    # reset the vminfo from status to db
    real_status_list = get_vms_status(cluster_ip)
    if isinstance(real_status_list, list):
        for i in real_status_list:
            i['Cluster'] = cluster_ip
        t_info_list = get_vms_info(cluster_ip)
        if isinstance(t_info_list, list):
            old_info = OldVmInfo.objects.filter(Cluster=cluster_ip)

            if reset:
                if type(real_status_list) is list and type(t_info_list) is list:
                    for i in t_info_list:
                        i.pop('id')
                # if different, clean the OldVmInfo in cluster_ip, then reset with the real_status_list
                if datadiff.diff(sorted(real_status_list), sorted(t_info_list)):
                    # clean OldVmInfo table
                    old_info_list = list(old_info.values())
                    old_info.delete()
                    # get OldVmInfo data
                    old_cluster_vms = VmInfo.objects.filter(Cluster=cluster_ip)
                    # reset the OldVmInfo table
                    for j in old_cluster_vms.values():
                        OldVmInfo.objects.get_or_create(**j)
                    # clean VmInfo table
                    old_cluster_vms.delete()
                    # reset the VmInfo table with real_status_list
                    for k in real_status_list:
                        VmInfo.objects.get_or_create(**k)
                    return {'now': real_status_list,
                            'old': t_info_list,
                            'deleted': old_info_list
                            }
                else:
                    return {'warning': 'The information is same, Nothing to do! - from reset_vms_info'}
            else:
                return {'warning': 'Reset is False, Nothing to do! - from reset_vms_info'}
        elif isinstance(t_info_list, dict):
            return t_info_list
    else:
        return {'error': 'The import Cluster IP not found in system - from reset_vms_info'}


def create_vms_info(cluster_ip):
    info_list = get_vms_status(cluster_ip)

    for i in info_list:
        i['Cluster'] = cluster_ip
        VmInfo.objects.get_or_create(**i)

    return {'succeed': 'Created!'}
