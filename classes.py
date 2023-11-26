class Standard:
    def __init__(self, standard_id, **kwargs):
        self.standard_id = standard_id
        self.total_answers = int
        self.right_answers = int
# should total and right answers be dictionary entries under the standard id?

    def add_to_list(self, standard_id):
        standard_list.append(standard_id)


class Student(Standard):
    # self.standard_list = standard_list

    def __init__(self, first_name, last_name, student_id, standard_id, total_answers, right_answers):
        self.first_name = first_name
        self.last_name = last_name
        self.student_id = student_id
        super().__init__(standard_id=standard_id,
                         total_answers=total_answers,
                         right_answers=right_answers)

    def __repr__(self):
        return f'{self.first_name}, {self.last_name}, {self.student_id}'


# Create list object for student names to exist globally ?
#   and subsequently standards and averages associated with student names
# Create list object for standards to exist globally ?
