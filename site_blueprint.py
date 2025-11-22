# ============================================================
# NetBox Job: Site Blueprint Generator
# Dynamically builds a plant site (PV / BESS / hybrid)
# from a parameterized YAML template.
# ============================================================

from extras.jobs import Job
from tenancy.models import Tenant
from dcim.models import Site, Device, DeviceType, DeviceRole, Manufacturer, Rack
from ipam.models import VLAN, Prefix, Namespace, IPAddress
from utilities.forms.fields import ChoiceField, IntegerField, CharField, BooleanField
from django.conf import settings
import yaml
import os


# ------------------------------------------------------------
# Helper: Load YAML Template
# ------------------------------------------------------------
def load_blueprint(template_name):
    template_dir = "/opt/netbox/site_templates"
    template_path = os.path.join(template_dir, f"{template_name}.yaml")

    with open(template_path, "r") as f:
        return yaml.safe_load(f)


# ------------------------------------------------------------
# The Script Class
# ------------------------------------------------------------
class DeploySiteFromBlueprint(Job):

    # ===== FORM INPUTS (High-Level Parameters) =====
    site_code = CharField(required=True, help_text="Example: S2025-001")
    plant_type = ChoiceField(
        choices=(
            ('PV', 'PV Plant'),
            ('BESS', 'Battery Energy Storage'),
            ('HYBRID', 'PV + BESS Hybrid'),
        ),
        required=True
    )
    field_switch_vendor = ChoiceField(
        choices=(
            ('Hirschmann', 'Hirschmann'),
            ('Moxa', 'Moxa'),
            ('Cisco', 'Cisco'),
        ),
        required=True
    )
    inverter_count = IntegerField(required=True, min_value=1)
    tracker_rows = IntegerField(required=True, min_value=0)
    met_station_count = IntegerField(required=True, min_value=0)

    network_base = CharField(required=True, help_text="Example: 10.100.")
    network_gateway = CharField(required=True, help_text="Example: 10.100.90.254")
    mgmt_vlan = IntegerField(required=True)

    enable_bess = BooleanField(required=False, initial=False)
    bess_container_count = IntegerField(required=False, initial=0)

    blueprint_template = CharField(
        required=True,
        initial="base_site_template",
        help_text="YAML template name (without .yaml)"
    )

    class Meta:
        name = "Deploy Site from Blueprint"
        description = "Creates an entire plant site from high-level parameters and a YAML blueprint."

    # ============================================================
    # Main Execution
    # ============================================================
    def run(self, data, commit=True):

        # ------------------------------------------------------------
        # Load / render the YAML blueprint
        # ------------------------------------------------------------
        self.log_info("Loading site blueprint template…")
        template = load_blueprint(data["blueprint_template"])

        # Merge user parameters into template context
        context = dict(data)
        context.update({"site_code": data["site_code"]})

        # ------------------------------------------------------------
        # 1. Create the Site
        # ------------------------------------------------------------
        site_name = data["site_code"]
        self.log_info(f"Creating site: {site_name}")

        site, created = Site.objects.get_or_create(
            name=site_name,
            defaults={
                "status": "active",
                "region": None,
                "description": f"Auto-generated site for {data['plant_type']} plant"
            }
        )
        if created:
            self.log_success(f"Site created: {site}")
        else:
            self.log_warning(f"Site already existed: {site}")

        # ------------------------------------------------------------
        # 2. Create VLANs
        # ------------------------------------------------------------
        self.log_info("Creating VLANs…")
        for vlan_def in template["vlans"]:
            vid = int(eval(str(vlan_def["vid"]), {}, context))  # supports {{ mgmt_vlan }}

            vlan_name = vlan_def["name"]
            vlan, _ = VLAN.objects.get_or_create(
                site=site,
                vid=vid,
                defaults={"name": vlan_name}
            )
            self.log_success(f"VLAN {vid} - {vlan_name}")

        # ------------------------------------------------------------
        # 3. Create Prefixes
        # ------------------------------------------------------------
        self.log_info("Creating prefixes…")

        prefix_map = template.get("prefixes", {})
        for key, cidr_tpl in prefix_map.items():
            cidr = cidr_tpl.format(**context)

            prefix, _ = Prefix.objects.get_or_create(
                prefix=cidr,
                defaults={"site": site}
            )
            self.log_success(f"Prefix {cidr}")

        # ------------------------------------------------------------
        # 4. Create Racks
        # ------------------------------------------------------------
        if "racks" in template:
            self.log_info("Creating racks…")

            core_rack_def = template["racks"].get("core_rack")
            if core_rack_def:
                rack, _ = Rack.objects.get_or_create(
                    name=core_rack_def["name"],
                    site=site,
                    defaults={"u_height": core_rack_def["height"]}
                )
                self.log_success(f"Rack created: {rack}")

        # ------------------------------------------------------------
        # 5. Create Devices: FIELD SWITCHES, INVERTERS, TRACKERS, MET STATIONS
        # ------------------------------------------------------------
        self._create_field_switches(site, template, context)
        self._create_inverters(site, template, context)
        self._create_trackers(site, template, context)
        self._create_met_stations(site, template, context)

        if data["enable_bess"]:
            self._create_bess_containers(site, template, context)

        self.log_success("Site deployment complete!")


    # ============================================================
    #  DEVICE CREATION HELPERS
    # ============================================================

    def _get_devicetype(self, model_name):
        try:
            return DeviceType.objects.get(model=model_name)
        except DeviceType.DoesNotExist:
            raise RuntimeError(f"DeviceType '{model_name}' not found in NetBox!")


    def _create_field_switches(self, site, template, context):
        vendor = context["field_switch_vendor"]
        vendor_map = template["switching"]["device_type_map"][vendor]

        count = (context["tracker_rows"] // template["devices"]["trackers"]["per_field_switch"]) or 1
        self.log_info(f"Creating {count} field switches for vendor {vendor}…")

        for i in range(1, count + 1):
            name = f"FSW-{i}"
            dt = self._get_devicetype(vendor_map["model"])

            Device.objects.get_or_create(
                name=name,
                site=site,
                device_type=dt,
                device_role=DeviceRole.objects.get_or_create(name="field-switch")[0]
            )
            self.log_success(f"Created field switch {name}")


    def _create_inverters(self, site, template, context):
        count = context["inverter_count"]
        inv_def = template["devices"]["inverters"]

        try:
            model_name = inv_def["device_type_map"][inv_def["vendor"]]
        except KeyError:
            raise RuntimeError(f"Inverter vendor {inv_def['vendor']} not mapped in template!")

        dt = self._get_devicetype(model_name)

        self.log_info(f"Creating {count} inverters…")
        for i in range(1, count + 1):
            name = f"INV-{i}"

            Device.objects.get_or_create(
                name=name,
                site=site,
                device_type=dt,
                device_role=DeviceRole.objects.get_or_create(name="inverter")[0]
            )
            self.log_success(f"Inverter {name}")


    def _create_trackers(self, site, template, context):
        count = context["tracker_rows"]

        dt = DeviceType.objects.get(model="Tracker-Controller")  # example
        role = DeviceRole.objects.get_or_create(name="tracker-controller")[0]

        self.log_info(f"Creating {count} tracker controllers…")
        for i in range(1, count + 1):
            name = f"TRK-{i}"
            Device.objects.get_or_create(
                name=name,
                site=site,
                device_type=dt,
                device_role=role
            )
            self.log_success(f"Tracker {name}")


    def _create_met_stations(self, site, template, context):
        count = context["met_station_count"]
        dt = DeviceType.objects.get(model="MET-Station")  # example
        role = DeviceRole.objects.get_or_create(name="met-station")[0]

        self.log_info(f"Creating {count} MET stations…")
        for i in range(1, count + 1):
            name = f"MET-{i}"
            Device.objects.get_or_create(
                name=name,
                site=site,
                device_type=dt,
                device_role=role
            )
            self.log_success(f"MET {name}")


    def _create_bess_containers(self, site, template, context):
        count = context["bess_container_count"]

        dt = DeviceType.objects.get(model="BESS-Container-Controller")
        role = DeviceRole.objects.get_or_create(name="bess-container")[0]

        self.log_info(f"Creating {count} BESS containers…")
        for i in range(1, count + 1):
            name = f"BESS-{i}"
            Device.objects.get_or_create(
                name=name,
                site=site,
                device_type=dt,
                device_role=role
            )
            self.log_success(f"BESS Container {name}")
