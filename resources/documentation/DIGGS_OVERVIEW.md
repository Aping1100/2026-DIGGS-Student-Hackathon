# DIGGS Technical Overview

## What is DIGGS?

DIGGS (Data Interchange for Geotechnical and Geoenvironmental Specialists) is an XML-based data transfer standard developed by the geotechnical community to facilitate the exchange of subsurface data.

## Key Components

### 1. XML/GML Structure
- Based on Geography Markup Language (GML)
- Extensible and machine-readable
- Self-documenting with metadata

### 2. Core Data Types

#### Project Data
- Project identification
- Location information
- Participants (client, contractor, engineer)
- Coordinate reference systems

#### Exploration Data
- Borehole descriptions (from boring logs)
- Test (trial) pit descriptions
- Sampling information
- Field observations

#### Testing Data
- Laboratory tests (grain size, Atterberg limits, strength tests, environmental tests)
- In-situ tests (SPT, CPT, pressuremeter)
- Material Tests

#### Interpreted Data
- Soil/rock stratigraphy
- Engineering properties
- Design parameters

## DIGGS Instance Structure

```xml
<diggs:Diggs>
  <diggs:project>
    <diggs:Project gml:id="p1">
    <gml:name>Project Xanadu</gml:name>
      <!-- Project metadata -->
    </diggs:Project>
  <diggs:project>
  
  <diggs:samplingFeature>
    <diggs:Borehole gml:id="bh1">
      <gml:name>Boring 1</gml:name>
      <diggs:projectRef xlink:href="#p1"/>
      <diggs:referencePoint>
        <diggs:PointLocation gml:id="pl-B-01" srsDimension="3"  srsName="https://www.opengis.net/def/crs-compound?1=http://www.opengis.net/def/crs/EPSG/0/4326%262=http://www.opengis.net/def/crs/EPSG/0/6360" uomLabels="dega dega ft">     
          <gml:pos>30.429139 -91.212861 19.00</gml:pos>
        </diggs:PointLocation>
      </diggs:referencePoint>
      <diggs:centerLine>
        <diggs:LinearExtent gml:id="cl-B-01" srsDimension="3" srsName="https://www.opengis.net/def/crs-compound?1=http://www.opengis.net/def/crs/EPSG/0/4326%262=http://www.opengis.net/def/crs/EPSG/0/6360" uomLabels="dega dega ft">
          <gml:posList>30.429139 -91.212861 19.00 30.429139 -91.212861 -141</gml:posList>
        </diggs:LinearExtent>
      </diggs:centerLine>
      <diggs:linearReferencing>
        <diggs:LinearSpatialReferenceSystem gml:id="cptsr1">
          <gml:identifier codeSpace="LADOT">LADOT:cptsr1</gml:identifier>     
          <glr:linearElement xlink:href="#ls1"/>
          <diggs:lrm xlink:href="http://diggsml.org/def/crs/DIGGS/0.1/lrm.xml#md_ft"/>
        </diggs:LinearSpatialReferenceSystem>
  `    </diggs:linearReferencing>   

      <!-- Other Borehole information -->
    </diggs:Borehole>
  </diggs:samplingFeature>

  <diggs:samplingActivity>
    <diggs:SamplingActivity gnl:id="sa1">
      <diggs:projectRef xlink:href="#p1"/>
      <diggs:samplingFeatureRef xlink:href="#bh1"/>

      <!-- Sampling Actifity info -->
    </diggs:SamplingActivity>
  <diggs:samplingActivity>
  
  <diggs:sample>
    <diggs:Sample gml:id="s1">
      <diggs:projectRef xlink:href="#p1"/>

      <!-- Sample details -->
    </diggs:Sample>
  </diggs:sample>
  
  <diggs:observation>
    <diggs:LithologySystem gml:id ="ls1">
      <diggs:lithologyObservation>
        <diggs:LithologyObservation gml:id="lo1">

        <!--Soil dewscription info for a single location -->
        </diggs:LithologyObservation>
      </diggs:lithologyObservation>
    </diggs:LithologySystem>
  </diggs:observation>
  
  <diggs:measurement>
    <diggs:Test gml:id="test1">
      <diggs:outcome>
        <diggs:TestResult>

        <!-- Lab or insitu test results -->
        </diggs:TestResult>
      </diggs:outcome>
      <diggs:procedure>

      <!-- Associated procedure object (eg. <diggs:AtterbergLimitsTest) -->
      </diggs:procedure>
    </diggs:Test>
  </diggs:measurement>
</diggs:Diggs>
```

## Common DIGGS Elements

### Spatial Elements
- `diggs:PointLocation` - Point coordinates (extends gml:Point)
- `diggs:LinearExtent` - Linear elements (extends gml:LineString)
- `diggs:PlanarSurface` - Area or surface elements (extends gml:Polygon)

### Measurement Elements
- `<diggs:totalMeasuredDepth uom="ft">400</diggs:totalMeasuredDepth>` - Length measurement
- `<diggs:bulkDensity uom="g/cm3">2.4</diggs:bulkDensity>` - Mass per volume measurement
- `<diggs:waterContent uom="%">24.7</diggs:waterContent>` - Force per force or dimensionless measurement

### Lithology Classification Elements
- `<classificationCode codeSpace="https://diggsml.org/def/codes/DIGGS/0.1/astmD2487.xml#gCHs">Gravelly fat clay with sand</classificationCode>` - Unified Soil Classification
  
- `<classificationCode codeSpace="https://diggsml.org/def/codes/DIGGS/0.1/aashtoM145.xml#a-1-b">A-1-b</classificationCode>` - AASHTO soil classification
  
- `<classificationSymbol codeSpace="https://diggsml.org/def/codes/DIGGS/0.1/grp-astmD2487.xml#CH">CH</classificationSymbol>` - Unified Soil Classification Group Symbol

## Working with DIGGS

### Parsing DIGGS Files
1. Handle XML namespaces properly
2. Validate against XSD schema
3. Extract relevant data elements
4. Handle missing/optional elements

### Creating DIGGS Files
1. Follow schema requirements
2. Include required metadata
3. Use appropriate units
4. Validate output

## Best Practices

1. **Always validate** DIGGS files against the schema
2. **Preserve precision** of numerical data
3. **Include units** for all measurements
4. **Document assumptions** in metadata
5. **Handle coordinates** carefully (check CRS)

## Common Challenges

### Namespace Management
DIGGS uses multiple namespaces (diggs, gml, xlink). Ensure your parser handles them correctly.

### Large Files
DIGGS files can be large. Consider streaming parsers for better performance.

### Coordinate Systems
Different projects may use different coordinate reference systems. Always check and transform if needed.

### Data Quality
Not all DIGGS files are complete. Build robust error handling.

## Resources

- **[DIGGS Official Documentation](https://diggsml.org/docs/)** - Complete specification and technical details
- **[DIGGS GitHub Repository](https://github.com/DIGGSml)** - Official schemas and resources
- **[Latest Schema Development](https://github.com/DIGGSml/schema-dev)** - Current DIGGS schema development
- **[DIGGS Tools & Validator](https://geosetta.org/web_map/map/DIGGS_Tools)** - Official validation and conversion tools
- [DIGGS Official Website](https://www.diggsml.org/)
- [GML Documentation](https://www.ogc.org/standards/gml)

## Important Tools

### DIGGS Validator
Use the official validator at [Geosetta DIGGS Tools](https://geosetta.org/web_map/map/DIGGS_Tools) to:
- Validate your DIGGS XML files
- Check schema compliance
- Convert between DIGGS versions
- Test your generated DIGGS output

---

*This overview provides the technical foundation for working with DIGGS data in the hackathon.*