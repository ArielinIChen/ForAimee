# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from django.db import models

# Create your models here.


class VmInfo(models.Model):
    Cluster = models.CharField(max_length=100)
    Name = models.CharField(max_length=100)
    Path = models.CharField(max_length=200)
    Guest = models.CharField(max_length=100)
    Annotation = models.CharField(max_length=100)
    powerStatus = models.CharField(max_length=50)
    instanceUuid = models.CharField(max_length=100)
    IP = models.CharField(max_length=50)

    def __unicode__(self):
        return self.Name


class OldVmInfo(models.Model):
    Cluster = models.CharField(max_length=100)
    Name = models.CharField(max_length=100)
    Path = models.CharField(max_length=200)
    Guest = models.CharField(max_length=100)
    Annotation = models.CharField(max_length=100)
    powerStatus = models.CharField(max_length=50)
    instanceUuid = models.CharField(max_length=100)
    IP = models.CharField(max_length=50)

    def __unicode__(self):
        return self.Name


class ClusterInfo(models.Model):
    Name = models.CharField(max_length=100)
    IP = models.GenericIPAddressField()
    User = models.CharField(max_length=50, default='root')
    Pwd = models.CharField(max_length=200, null=True)

    def __unicode__(self):
        return self.Name


class OldClusterInfo(models.Model):
    Name = models.CharField(max_length=100)
    IP = models.GenericIPAddressField()
    User = models.CharField(max_length=50, default='root')
    Pwd = models.CharField(max_length=200, null=True)

    def __unicode__(self):
        return self.Name


class VCenterAliveStatus(models.Model):
    Name = models.CharField(max_length=100)
    LastUpdateTime = models.DateTimeField(auto_now=True)
    LastUpdateStatus = models.CharField(max_length=20, default='Unknown')

    def __unicode__(self):
        return self.Name


class VCenterInfo(models.Model):
    IP = models.GenericIPAddressField()
    Authorization = models.CharField(max_length=200)
    CheckUrl = models.CharField(max_length=200)
    Status = models.ForeignKey(VCenterAliveStatus)

    def __unicode__(self):
        return self.Status.Name
