public final class Employee {
    private String firstName;
    private String lastName;
    private String nationality;


    public Employee(String firstName, String lastName, String nationality) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.nationality = nationality;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public String getNationalID() {
        return nationality;
    }

    @Override
    public String toString() {
        return "first_name:" + firstName + '\t' +
               "last_name:" + lastName + '\t' +
               "nationality:" + nationality;
    }
}
