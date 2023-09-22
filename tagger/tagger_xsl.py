import GoogleSheet as gs
import boto3
import boto3.session
import json


class awsResources:
    def __init__(self, profile, type):
        self.session = boto3.Session(profile_name=profile)
        self.resource = self.session.resource(type)


def get_aws_resources(session, tags, cloud):
    instances = []
    ec2 = session.resource("ec2")
    for i in ec2.instances.all():
        instance = [
            i.id,
            i.private_ip_address,
            cloud,
            i.key_name,
            i.state["Name"],
            i.instance_type,
            i.placement["AvailabilityZone"],
            "",
        ]
        basecount = len(instance)
        try:
            for key in tags[basecount:]:
                value = ""
                for tag in i.tags:

                    if tag["Key"] == key:

                        value = tag["Value"]
                instance.append(value)
        except Exception as e:
            print(e)
        instances.append(instance)
    return instances


def search(item, values):
    for x in values:
        if item[0] == x[0]:
            return values.index(x)
    return None


def main():

    with open("Files/Config.json") as config_file:
        Config = json.load(config_file)

    profiles = Config["profiles"]
    spredsheet_id = Config["spredsheet_id"]
    headers = Config["headers"]

    for profile in profiles:
        print(f"------------{profile}------------")
        try:
            session = boto3.Session(profile_name=profile)
        except Exception as e:
            print(e)
            continue
        sheet = gs.GoogleSheet("Master", id=spredsheet_id, baserange=profile)
        instances = get_aws_resources(session, headers, profile)

        print(f"https://docs.google.com/spreadsheets/d/{sheet.get_id()}/edit#gid=0")
        sheet.clear_sheet(headers, instances)
        sheet.sheet_write(headers, instances)

if __name__ == "__main__":
    main()
