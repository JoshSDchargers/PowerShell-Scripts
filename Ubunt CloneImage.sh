#!/bin/bash

#CHANGE DIRECTORY#
cd media/ubuntu/Win10x64

#CLONE IMAGE#
sudo dd bs=1M if=/dev/mmcblk0 | gzip>intel.img.gz

#DEPLOY IMAGE#
zcat intel.img.gz | sudo dd of=/dev/mmcblk0 bs=1M