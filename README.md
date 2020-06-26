# exports-java
Export objects to varius files. Excel, csv and etc

# Maven dependency:
```
<dependency>
  <groupId>com.github.telmedola</groupId>
  <artifactId>exports-java</artifactId>
  <version>1.0.0</version>
</dependency>
```
# Gradle dependency:
```
dependencies {
    implementation 'com.github.telmedola:exports-java:1.0.0'
  }
```

# Example of use:
```
public class TestExport {
	
public static void main(String[] args) throws IllegalAccessException, IOException, ParseException {

		List<Test> lista = new ArrayList();

		Test a;
		for (int i = 0; i<3; i++){
			a = new Test();
			a.setId(Long.parseLong(Integer.toString(i)));
			a.setName("Name "+ Integer.toString(i));
			a.setInclude(new Date());
			if (i > 1)
			    a.setActive("Y");
			else
			    a.setActive("N");
			a.setJoin("Not appear");

			lista.add(a);
		}

		Exports<Test> export = new Exports<>();

		System.out.println(export.exportListToCSV(lista));

		System.out.println(export.exportListToExcel(lista,"Customers"));
	}
}

class Test {
    private Long id;
    @ExportsName(name = "name of costumer")
    private String name;
    @JsonFormat(pattern = "yyyy-MM-dd")
    private Date include;
    @ExportsFromTo(from={"Y","N"}, to={"Yes","No"})
    private String active;
    @JsonIgnore
    private String join;
    @ExportsIgnore
    private String notExport;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getInclude() {
        return include;
    }

    public void setInclude(Date include) {
        this.include = include;
    }
    public String getActive() {
        return active;
    }

    public void setActive(String active) {
        this.active = active;
    }

    public String getJoin() {
        return join;
    }

    public void setJoin(String join) {
        this.join = join;
    }

}
```
