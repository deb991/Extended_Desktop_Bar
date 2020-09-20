import winreg
import subprocess
#from __builtin__ import file


def foo(hive, flag):
    aReg = winreg.ConnectRegistry(None, hive)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                          0, winreg.KEY_READ | flag)

    count_subkey = winreg.QueryInfoKey(aKey)[0]

    software_list = []

    for i in range(count_subkey):
        software = {}
        try:
            asubkey_name = winreg.EnumKey(aKey, i)
            asubkey = winreg.OpenKey(aKey, asubkey_name)
            software['name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]

            try:
                software['version'] = winreg.QueryValueEx(asubkey, "DisplayVersion")[0]
            except EnvironmentError:
                software['version'] = 'undefined'
            try:
                software['publisher'] = winreg.QueryValueEx(asubkey, "Publisher")[0]
            except EnvironmentError:
                software['publisher'] = 'undefined'
            software_list.append(software)
        except EnvironmentError:
            continue

    return software_list

software_list = foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_64KEY) + foo(winreg.HKEY_CURRENT_USER, 0)

for software in software_list:
    print('Name=%s, Version=%s, Publisher=%s' % (software['name'], software['version'], software['publisher']))
print('Number of installed apps: %s' % len(software_list))

try:
    p = subprocess.Popen(args, stdout=subprocess.PIPE, shell=True)
    (stdout, stderr) = p.communicate()
    with open('pkg/xml/appList.xml', 'w') as fp:
        fp.write(stdout)

except "TypeError: __init__() missing 1 required positional argument: 'args'":
    p = subprocess.Popen(args, stdout=subprocess.PIPE, shell=True)
    (stdout, stderr) = p.communicate()
    with open('pkg/xml/appList.xml', 'w') as fp:
        fp.write(stdout)