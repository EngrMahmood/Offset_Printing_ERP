from typing import Optional


class OEECalculator:
    @staticmethod
    def availability(run_time: float, downtime_minutes: float) -> float:
        total_available = float(run_time or 0) + float(downtime_minutes or 0)
        if total_available == 0:
            return 0.0
        return float(run_time or 0) / total_available

    @staticmethod
    def press_utilization(make_ready_time: float, run_time: float, downtime_minutes: float) -> float:
        total_session = float(make_ready_time or 0) + float(run_time or 0) + float(downtime_minutes or 0)
        if total_session == 0:
            return 0.0
        return float(run_time or 0) / total_session

    @staticmethod
    def performance(impressions: float, standard_impressions_per_hour: float, run_time: float) -> float:
        if not standard_impressions_per_hour or not run_time:
            return 0.0
        expected_impressions = float(standard_impressions_per_hour) * (float(run_time) / 60.0)
        if expected_impressions == 0:
            return 0.0
        return float(impressions or 0) / expected_impressions

    @staticmethod
    def quality(output_sheets: float, waste_sheets: float, waste_reason: Optional[str]) -> float:
        if output_sheets is None:
            output_sheets = 0
        if waste_sheets is None:
            waste_sheets = 0
        quality_affecting = {'paper_jam', 'color_issue', 'operator_error', 'machine_issue', 'other'}
        quality_waste = float(waste_sheets or 0) if waste_reason in quality_affecting else 0.0
        good_sheets = float(output_sheets or 0)
        total_quality = good_sheets + quality_waste
        if total_quality == 0:
            return 0.0
        return good_sheets / total_quality

    @staticmethod
    def oee(availability: float, performance: float, quality: float) -> float:
        return round(float(availability or 0) * float(performance or 0) * float(quality or 0), 2)
