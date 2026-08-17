from rest_framework import serializers

MIN_TOTAL_TEAM_SIZE = 1
MAX_TOTAL_TEAM_SIZE = 5
DIFFICULTY_CHOICES = ["Easy", "Medium", "Hard"]

from . import excel_utils


class MemberSerializer(serializers.Serializer):
    name = serializers.CharField(max_length=150)
    roll_no = serializers.CharField(max_length=50, required=False, allow_blank=True)
    email = serializers.EmailField(required=False, allow_blank=True)
    phone = serializers.CharField(max_length=20, required=False, allow_blank=True)


class RegistrationSerializer(serializers.Serializer):
    team_name = serializers.CharField(max_length=150)
    track = serializers.CharField(max_length=100)
    college = serializers.CharField(max_length=200)

    leader_name = serializers.CharField(max_length=150)
    leader_roll_no = serializers.CharField(max_length=50)
    leader_email = serializers.EmailField()
    leader_phone = serializers.CharField(max_length=20)

    idea = serializers.CharField(max_length=2000, required=False, allow_blank=True)

    members = MemberSerializer(many=True, required=False)
    
    
    def validate_leader_roll_no(self, value):
        if excel_utils.find_team_by_roll_no(value) is not None:
            raise serializers.ValidationError(
                "A team is already registered with this leader roll number."
            )
        return value

    def validate_members(self, members):
        total_size = 1 + len(members)
        if total_size < MIN_TOTAL_TEAM_SIZE:
            raise serializers.ValidationError(
                f"Team must have at least {MIN_TOTAL_TEAM_SIZE} members including the leader."
            )
        if total_size > MAX_TOTAL_TEAM_SIZE:
            raise serializers.ValidationError(
                f"Team can have at most {MAX_TOTAL_TEAM_SIZE} members including the leader."
            )
        return members




class TeamLoginSerializer(serializers.Serializer):
    """Team leader logs in with just their roll number."""
    roll_no = serializers.CharField(max_length=50)


class ChooseProblemSerializer(serializers.Serializer):
    """Team leader picks (or changes) their team's single problem statement."""
    problem_id = serializers.IntegerField()


class TeamIdeaSerializer(serializers.Serializer):
    """Team leader saves an idea / abstract for their team."""
    idea = serializers.CharField(max_length=2000, required=True, allow_blank=True)


class ProblemStatementSerializer(serializers.Serializer):
    """Used for CREATE — title is required."""
    title = serializers.CharField(max_length=200)
    description = serializers.CharField(max_length=4000, required=False, allow_blank=True)
    track = serializers.CharField(max_length=100, required=False, allow_blank=True)
    difficulty = serializers.ChoiceField(
        choices=DIFFICULTY_CHOICES, required=False, allow_blank=True
    )


class ProblemStatementUpdateSerializer(ProblemStatementSerializer):
    """Used for UPDATE — every field optional so PATCH can send just one."""
    title = serializers.CharField(max_length=200, required=False)