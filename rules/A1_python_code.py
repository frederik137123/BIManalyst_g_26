
import ifcopenshell

# Open the IFC model
# model = ifcopenshell.open('C:/Users/friis/OneDrive - Danmarks Tekniske Universitet/DTU/41934 - Advanced BIM/IFC/CES_BLD_24_06_STR.ifc')

# Function to find property value by name
def get_property_value(property_set, property_name):
    for prop in property_set.HasProperties:
        if prop.Name == property_name:
            return prop.NominalValue.wrappedValue
    return None

# Get all beams in the model
beams = model.by_type('IfcBeam')
print(f"Beams in model: {len(beams)}")

# Initialize a variable to track the maximum span
max_span = 0
max_span_beam_id = None

# Loop through each beam to get the span from the property sets
for beam in beams:
    span_found = False

    # Check property sets for span
    if hasattr(beam, "IsDefinedBy"):
        for definition in beam.IsDefinedBy:
            if definition.is_a("IfcRelDefinesByProperties"):
                prop_set = definition.RelatingPropertyDefinition
                if prop_set.is_a("IfcPropertySet"):
                    # Check for Span property
                    span = get_property_value(prop_set, "Span")
                    if span is not None:
                        # Update maximum span if the current span is larger
                        if span > max_span:
                            max_span = span
                            max_span_beam_id = beam.Name
                        span_found = True
                        break

# Print the maximum span and the beam that has it
if max_span_beam_id:
    print(f"Maximum span: {max_span} mm, found in beam: {max_span_beam_id}")
else:
    print("No spans found in the property sets of the beams.")

# The GlobalId of the specific element you want to find
#target_global_id = max_span_beam_id

# Find the element by GlobalId
#element = model.by_guid(target_global_id)

#if element:
 #   print(f"Element found: {element}")
  #  print(f"Element type: {element.is_a()}")
#else:
   # print(f"No element found with GlobalId: {target_global_id}")