# coding: utf-8
import os
import sys
import contextlib
from sqlalchemy import (Column, ForeignKey,
    Integer, String, Boolean, UniqueConstraint, Date)
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from sqlalchemy.sql import default_comparator
from sqlalchemy import create_engine
 
Base = declarative_base()

# string parsing constants
STATE_PREFIX = 0
STATE_NUMBER = 1
STATE_SEPARATOR = 2
STATE_SUFFIX = 3

def _check_min_max(num, min, max):
    new_min = min if min is not None and num >= min else num
    new_max = max if max is not None and num <= max else num
    return new_min, new_max

def _check_prev(cur, prev):
    return cur, cur < prev

def _check_year_in_interval(year, min, max):
    return (min is None or year >= min) and (max is None or year <= max)

def _get_min_year(issues):
    return min(issue.year for issue in issues)

def _get_max_year(issues):
    return max(issue.year for issue in issues)

def _get_min_number(issues):
    return min(issue.num_min for issue in issues if issue.num_min is not None)

def _get_max_number(issues):
    return max(issue.num_max for issue in issues if issue.num_max is not None)

def _contains_issue(issues, year, number):
    return next((True for issue in issues if issue.num_min is not None and issue.num_max is not None
        and issue.num_min <= number <= issue.num_max and (year is None or issue.year == year)), False)

def extract_issue_numbers(s):
    if len(s) == 0:
        return None, None, None
    state = STATE_PREFIX
    cur, prev = 0, 0
    min, max = None, None
    inv = False
    suffix = ''
    for c in s:
        if state == STATE_SUFFIX:
            suffix += c
        elif c.isdigit():
            if state == STATE_PREFIX or state == STATE_SEPARATOR:
                state = STATE_NUMBER
                cur = int(c)
            elif state == STATE_NUMBER:
                cur = cur*10 + int(c)
        elif c == '/' or c == '-':
            if state == STATE_NUMBER:
                state = STATE_SEPARATOR
                min, max = _check_min_max(cur, min, max)
                prev, inv = _check_prev(cur, prev)
        else:
            if state == STATE_NUMBER:
                state = STATE_SUFFIX
                suffix = c
                min, max = _check_min_max(cur, min, max)
                prev, inv = _check_prev(cur, prev)

    if state == STATE_NUMBER:
        min, max = _check_min_max(cur, min, max)
        prev, inv = _check_prev(cur, prev)

    return min, max, inv, suffix.strip()

 
class Magazine(Base):
    __tablename__ = 'magazine'
    id = Column(Integer, primary_key=True)
    name = Column(String(250), nullable=False, unique=True)

    show_columns = ('name',)
    column_names = ('nome',)
    sort_columns = (True,)
    sort_fields = ('name',)
    joined_tables=(None,)
    filter_columns = (True,)
    edit_columns = (False,)

    __table_args__ = (
        UniqueConstraint("name"),
    )

    def __repr__(self):
        return "<Magazine(name='{}')>".format(self.name)

    def get_current_issues_for_numbering(self, numbering):
        from_year = None if numbering.from_year is None else numbering.from_year
        to_year = None if numbering.to_year is None else numbering.to_year
        return [issue for issue in self.issues if _check_year_in_interval(issue.year, from_year, to_year)]

    def get_all_issues_for_numbering(self, numbering, current_issues=None):
        if current_issues is None:
            current_issues = self.get_current_issues_for_numbering(numbering)

        if numbering.is_yearly:
            from_year = _get_min_year(current_issues) if numbering.from_year is None else numbering.from_year
            to_year = _get_max_year(current_issues) if numbering.to_year is None else numbering.to_year
            from_number = 1 if numbering.from_number is None else numbering.from_number
            to_number = 12 if numbering.to_number is None else numbering.to_number
            return [(year, number) for year in range(from_year, to_year+1) for number in range(from_number, to_number+1)]
        
        from_number = _get_min_number(current_issues) if numbering.from_number is None else numbering.from_number
        to_number = _get_max_number(current_issues) if numbering.to_number is None else numbering.to_number
        return [(None, number) for number in range(from_number, to_number+1)]

    def _get_missing_numbers(self, numbering):
        current_issues = self.get_current_issues_for_numbering(numbering)
        all_numbers = self.get_all_issues_for_numbering(numbering, current_issues)
        return [number for number in all_numbers if not _contains_issue(current_issues, number[0], number[1])]

    def get_missing_numbers(self):
        missing_numbers = []
        for numbering in self.numberings:
            missing_numbers += self._get_missing_numbers(numbering)
        return missing_numbers
        
 
class Issue(Base):
    __tablename__ = 'issue'
    id = Column(Integer, primary_key=True)
    year = Column(Integer)
    issue_number = Column(String(250), nullable=False)
    copies = Column(Integer, nullable=False)
    is_new = Column(Boolean, nullable=False)
    magazine_id = Column(Integer, ForeignKey('magazine.id'))
    num_min = Column(Integer)
    num_max = Column(Integer)
    inv = Column(Boolean)
    suffix = Column(String(250))
    magazine = relationship("Magazine", back_populates="issues")

    show_columns = ('magazine.name', 'year', 'issue_number', 'copies', 'is_new')
    column_names = ('testata', 'anno', 'numero', 'copie', 'nuovo')
    sort_columns = (True, True, True, True, True)
    sort_fields = (Magazine.name, year, issue_number, copies, is_new)
    joined_tables = (Magazine, None, None, None, None)
    filter_columns = (True, True, True, True, True)
    edit_columns = (False, True, True, True, False)

    __table_args__ = (
        UniqueConstraint("magazine_id", "year", "issue_number"),
    )

    def populate_issue_numbers(self):
        self.num_min, self.num_max, self.inv, self.suffix = extract_issue_numbers(str(self.issue_number))

    def __repr__(self):
        return "<Issue(magazine_id={}, year={}, issue_number='{}', copies={}, is_new={})>"\
        .format(self.magazine_id, self.year, self.issue_number, self.copies, self.is_new)


class Numbering(Base):
    __tablename__ = 'numbering'
    id = Column(Integer, primary_key=True)
    magazine_id = Column(Integer, ForeignKey('magazine.id'))
    from_year = Column(Integer)
    to_year = Column(Integer)
    is_yearly = Column(Boolean, nullable=False)
    from_number = Column(Integer)
    to_number = Column(Integer)
    magazine = relationship("Magazine", back_populates="numberings")

    show_columns = ('magazine.name', 'from_year', 'to_year', 'is_yearly', 'from_number', 'to_number')
    column_names = ('testata', 'data inizio', 'data fine', 'annuale', 'da', 'a')
    sort_columns = (True, True, True, True, False, False)
    sort_fields = (Magazine.name, from_year, to_year, is_yearly, from_number, to_number)
    joined_tables = (Magazine, None, None, None, None, None)
    filter_columns = (True, True, True, True, True, True)
    edit_columns = (False, True, True, True, True, True)

    def __repr__(self):
        return "<Numbering(magazine_id={}, from_year={}, to_year={}, is_yearly={}, from_number={}, to_number={})>"\
        .format(self.magazine_id, self.from_year, self.to_year, self.is_yearly, self.from_number, self.to_number)


Magazine.issues = relationship("Issue", order_by=Issue.id, back_populates="magazine")
Magazine.numberings = relationship("Numbering", order_by=Numbering.id, back_populates="magazine")
