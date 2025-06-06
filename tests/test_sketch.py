import os

from cadquery.sketch import Sketch, Vector, Location
from cadquery.selectors import LengthNthSelector
from cadquery import Edge, Vertex

from pytest import approx, raises, fixture
from math import pi, sqrt

testdataDir = os.path.join(os.path.dirname(__file__), "testdata")


def test_face_interface():

    s1 = Sketch().rect(1, 2, 45)

    assert s1._faces.Area() == approx(2)
    assert s1.vertices(">X")._selection[0].toTuple()[0] == approx(1.5 / sqrt(2))

    s2 = Sketch().circle(1)

    assert s2._faces.Area() == approx(pi)

    s3 = Sketch().ellipse(2, 0.5)

    assert s3._faces.Area() == approx(pi)

    s4 = Sketch().trapezoid(2, 0.5, 45)

    assert s4._faces.Area() == approx(0.75)

    s4 = Sketch().trapezoid(2, 0.5, 45)

    assert s4._faces.Area() == approx(0.75)

    s5 = Sketch().slot(3, 2)

    assert s5._faces.Area() == approx(6 + pi)
    assert s5.edges(">Y")._selection[0].Length() == approx(3)

    s6 = Sketch().regularPolygon(1, 5)

    assert len(s6.vertices()._selection) == 5
    assert s6.vertices(">Y")._selection[0].toTuple()[1] == approx(1)

    s7 = Sketch().polygon([(0, 0), (0, 1), (1, 0)])

    assert len(s7.vertices()._selection) == 3
    assert s7._faces.Area() == approx(0.5)

    s8 = Sketch().face(Sketch().rect(1, 1).val())
    assert s8._faces.Area() == approx(1)

    s9 = Sketch().face(Sketch().rect(1, 1))
    assert s9._faces.Area() == approx(1)

    with raises(ValueError):
        Sketch().face(1234)


def test_modes():

    # additive mode
    s1 = Sketch().rect(2, 2).rect(1, 1, mode="a")

    assert s1._faces.Area() == approx(4)
    assert len(s1._faces.Faces()) == 2

    # subtraction mode
    s2 = Sketch().rect(2, 2).rect(1, 1, mode="s")

    assert s2._faces.Area() == approx(4 - 1)
    assert len(s2._faces.Faces()) == 1

    # intersection mode
    s3 = Sketch().rect(2, 2).rect(1, 1, mode="i")

    assert s3._faces.Area() == approx(1)
    assert len(s3._faces.Faces()) == 1

    # construction mode
    s4 = Sketch().rect(2, 2).rect(1, 1, mode="c", tag="t")

    assert s4._faces.Area() == approx(4)
    assert len(s4._faces.Faces()) == 1
    assert s4._tags["t"][0].Area() == approx(1)

    # construction mode requires tagging
    with raises(ValueError):
        Sketch().rect(2, 2).rect(1, 1, mode="c")

    with raises(ValueError):
        Sketch().rect(2, 2).rect(1, 1, mode="dummy")

    # replace mode
    s5 = Sketch().rect(1, 1).wires().offset(-0.1, mode="r").reset()
    assert s5.val().Area() == approx(0.8 ** 2)


def test_distribute():

    with raises(ValueError):
        Sketch().rect(2, 2).faces().distribute(5)

    with raises(ValueError):
        Sketch().rect(2, 2).distribute(5)

    with raises(ValueError):
        Sketch().circle(1).wires().distribute(0, 0, 1)

    s1 = Sketch().circle(4, mode="c", tag="c").edges(tag="c").distribute(3)

    assert len(s1._selection) == approx(3)

    s1.rect(1, 1)

    assert s1._faces.Area() == approx(3)
    assert len(s1._faces.Faces()) == 3
    assert len(s1.reset().vertices("<X")._selection) == 2

    for f in s1._faces.Faces():
        assert f.Center().Length == approx(4)

    s2 = (
        Sketch()
        .circle(4, mode="c", tag="c")
        .edges(tag="c")
        .distribute(3, rotate=False)
        .rect(1, 1)
    )

    assert s2._faces.Area() == approx(3)
    assert len(s2._faces.Faces()) == 3
    assert len(s2.reset().vertices("<X")._selection) == 4

    for f in s2._faces.Faces():
        assert f.Center().Length == approx(4)

    s3 = (
        Sketch().circle(4, mode="c", tag="c").edges(tag="c").distribute(3, 0.625, 0.875)
    )

    assert len(s3._selection) == approx(3)

    s3.rect(1, 0.5).reset().vertices("<X")

    assert s3._selection[0].toTuple() == approx(
        (-3.358757210636101, -3.005203820042827, 0.0)
    )

    s3.reset().vertices(">X")

    assert s3._selection[0].toTuple() == approx(
        (3.358757210636101, -3.005203820042827, 0.0)
    )

    s4 = Sketch().arc((0, 0), 4, 180, 180).edges().distribute(3, 0.25, 0.75)

    assert len(s4._selection) == approx(3)

    s4.rect(1, 0.5).reset().faces("<X").vertices("<X")

    assert s4._selection[0].toTuple() == approx(
        (-3.358757210636101, -3.005203820042827, 0.0)
    )

    s4.reset().faces(">X").vertices(">X")

    assert s4._selection[0].toTuple() == approx(
        (3.358757210636101, -3.005203820042827, 0.0)
    )

    s5 = (
        Sketch()
        .arc((0, 2), 4, 0, 90)
        .arc((0, -2), 4, 0, -90)
        .edges()
        .distribute(4, 0, 1)
        .circle(0.5)
    )

    assert len(s5._selection) == approx(8)

    s5.reset().faces(">X").faces(">Y")

    assert s5._selection[0].Center().toTuple() == approx((4.0, 2.0, 0.0))

    s5.reset().faces(">X").faces("<Y")

    assert s5._selection[0].Center().toTuple() == approx((4.0, -2.0, 0.0))

    s5.reset().faces(">Y")

    assert s5._selection[0].Center().toTuple() == approx((0.0, 6.0, 0.0))

    # make sure that we can use distribute on straight lines
    _ = Sketch().segment((0, 0), (10, 0)).edges().distribute(3).rect(1, 2)


def test_rarray():

    with raises(ValueError):
        Sketch().rarray(2, 2, 3, 0).rect(1, 1)

    s1 = Sketch().rarray(2, 2, 3, 3).rect(1, 1)

    assert s1._faces.Area() == approx(9)
    assert len(s1._faces.Faces()) == 9

    s2 = Sketch().push([(0, 0), (1, 1)]).rarray(2, 2, 3, 3).rect(0.5, 0.5)

    assert s2._faces.Area() == approx(18 * 0.25)
    assert len(s2._faces.Faces()) == 18
    assert s2.reset().vertices(">(1,1,0)")._selection[0].toTuple() == approx(
        (3.25, 3.25, 0)
    )


def test_parray():

    with raises(ValueError):
        Sketch().parray(2, 0, 90, 0).rect(1, 1)

    s1 = Sketch().parray(2, 0, 90, 3).rect(1, 1)

    assert s1._faces.Area() == approx(3)
    assert len(s1._faces.Faces()) == 3

    s2 = Sketch().push([(0, 0), (1, 1)]).parray(2, 0, 90, 3).rect(0.5, 0.5)

    assert s2._faces.Area() == approx(6 * 0.25)
    assert len(s2._faces.Faces()) == 6

    s3 = Sketch().parray(2, 0, 90, 3, False).rect(0.5, 0.5).reset().vertices(">(1,1,0)")

    assert len(s3._selection) == 1
    assert s3._selection[0].toTuple() == approx(
        (1.6642135623730951, 1.664213562373095, 0.0)
    )

    s4 = Sketch().push([(0, 0), (0, 1)]).parray(2, 0, 90, 3).rect(0.5, 0.5)
    s4.reset().faces(">(0,1,0)")

    assert s4._selection[0].Center().Length == approx(3)

    s5 = Sketch().push([(0, 1)], tag="loc")

    assert len(s5._tags["loc"]) == 1

    s6 = Sketch().push([(-4, 1), (0, 0), (4, -1)]).parray(2, 10, 50, 3).rect(1.0, 0.5)
    s6.reset().vertices(">(-1,0,0)")

    assert s6._selection[0].toTuple() == approx(
        (-3.46650635094611, 2.424038105676658, 0.0)
    )

    s6.reset().vertices(">(1,0,0)")

    assert s6._selection[0].toTuple() == approx(
        (6.505431426947252, -0.8120814940857262, 0.0)
    )

    s7 = Sketch().parray(1, 135, 0, 1).circle(0.1)
    s7.reset().faces()

    assert len(s7._selection) == 1
    assert s7._selection[0].Center().toTuple() == approx(
        (-0.7071067811865475, 0.7071067811865476, 0.0)
    )

    s8 = Sketch().parray(4, 20, 360, 6).rect(1.0, 0.5)

    assert len(s8._faces.Faces()) == 6

    s8.reset().vertices(">(0,-1,0)")

    assert s8._selection[0].toTuple() == approx(
        (-0.5352148612481344, -4.475046932971669, 0.0)
    )

    s9 = (
        Sketch()
        .push([(-4, 1)])
        .circle(0.1)
        .reset()
        .faces()
        .parray(2, 10, 50, 3)
        .rect(1.0, 0.5, 40, "a", "rects")
    )

    assert len(s9._faces.Faces()) == 4

    s9.reset().vertices(">(-1,0,0)", tag="rects")

    assert s9._selection[0].toTuple() == approx(
        (-3.3330260270865173, 3.1810426396582487, 0.0)
    )


def test_each():

    s1 = Sketch().each(lambda l: Sketch().push([l]).rect(1, 1))

    assert len(s1._faces.Faces()) == 1

    s2 = (
        Sketch()
        .push([(0, 0), (2, 2)])
        .each(lambda l: Sketch().push([l]).rect(1, 1), ignore_selection=True)
    )

    assert len(s2._faces.Faces()) == 1


def test_modifiers():

    s1 = Sketch().push([(-2, 0), (2, 0)]).rect(1, 1).reset().vertices("<X").fillet(0.1)

    assert len(s1._faces.Faces()) == 2
    assert len(s1._faces.Edges()) == 10

    s2 = Sketch().push([(-2, 0), (2, 0)]).rect(1, 1).reset().vertices(">X").chamfer(0.1)

    assert len(s2._faces.Faces()) == 2
    assert len(s2._faces.Edges()) == 10

    s3 = Sketch().push([(-2, 0), (2, 0)]).rect(1, 1).reset().hull()

    assert len(s3._faces.Faces()) == 3
    assert s3._faces.Area() == approx(5)

    s4 = Sketch().push([(-2, 0), (2, 0)]).rect(1, 1).reset().hull()

    assert len(s4._faces.Faces()) == 3
    assert s4._faces.Area() == approx(5)

    s5 = (
        Sketch()
        .push([(-2, 0), (0, 0), (2, 0)])
        .rect(1, 1)
        .reset()
        .faces("not >X")
        .edges()
        .hull()
    )

    assert len(s5._faces.Faces()) == 4
    assert s5._faces.Area() == approx(4)

    s6 = Sketch().segment((0, 0), (0, 1)).segment((1, 0), (2, 0)).hull()

    assert len(s6._faces.Faces()) == 1
    assert s6._faces.Area() == approx(1)

    with raises(ValueError):
        Sketch().rect(1, 1).vertices().hull()

    with raises(ValueError):
        Sketch().hull()

    s7 = Sketch().rect(2, 2).wires().offset(1)

    assert len(s7._faces.Faces()) == 2
    assert len(s7._faces.Edges()) == 4 + 4 + 4

    s7.clean()

    assert len(s7._faces.Faces()) == 1
    assert len(s7._faces.Edges()) == 4 + 4

    s8 = Sketch().rect(2, 2).wires().offset(-0.5, mode="s")

    assert len(s8._faces.Faces()) == 1
    assert len(s8._faces.Edges()) == 4 + 4


def test_delete():

    s1 = Sketch().push([(-2, 0), (2, 0)]).rect(1, 1).reset()

    assert len(s1._faces.Faces()) == 2

    s1.faces("<X").delete()

    assert len(s1._faces.Faces()) == 1

    s2 = Sketch().segment((0, 0), (1, 0)).segment((0, 1), tag="e").close()
    assert len(s2._edges) == 3

    s2.edges("<X").delete()

    assert len(s2._edges) == 2

    with raises(ValueError):
        s2.vertices().delete()


def test_selectors():

    s = Sketch().push([(-2, 0), (2, 0)]).rect(1, 1).rect(0.5, 0.5, mode="s").reset()

    assert s._selection is None

    s.vertices()

    assert len(s._selection) == 16

    s.reset()

    assert s._selection is None

    s.edges()

    assert len(s._selection) == 16

    s.reset().wires()

    assert len(s._selection) == 4

    s.reset().faces()

    assert len(s._selection) == 2

    s.reset().vertices("<Y")

    assert len(s._selection) == 4

    s.reset().edges("<X or >X")

    assert len(s._selection) == 2

    s.tag("test").reset()

    assert s._selection is None

    s.select("test")

    assert len(s._selection) == 2

    s.reset().wires()

    assert len(s._selection) == 4

    s.reset().wires(LengthNthSelector(1))

    assert len(s._selection) == 2
    assert len(s.vals()) == 2

    s.reset().vertices("<X and <Y").val()
    assert s.val().toTuple() == approx((-2.5, -0.5, 0.0))

    with raises(IndexError):
        s.reset().vertices(">>X[1] and <Y").val()
    assert len(s._selection) == 0


def test_edge_interface():

    s1 = (
        Sketch()
        .segment((0, 0), (1, 0))
        .segment((1, 1))
        .segment(1, 180)
        .close()
        .assemble()
    )

    assert len(s1._faces.Faces()) == 1
    assert s1._faces.Area() == approx(1)

    s2 = Sketch().arc((0, 0), (1, 1), (0, 2)).close().assemble()

    assert len(s2._faces.Faces()) == 1
    assert s2._faces.Area() == approx(pi / 2)

    s3 = Sketch().arc((0, 0), (1, 1), (0, 2)).arc((-1, 1), (0, 0)).assemble()

    assert len(s3._faces.Faces()) == 1
    assert s3._faces.Area() == approx(pi)

    s4 = Sketch().arc((0, 0), 1, 0, 90)

    assert len(s4.vertices()._selection) == 2
    assert s4.vertices(">Y")._selection[0].Center().y == approx(1)

    s5 = Sketch().arc((0, 0), 1, 0, -90)

    assert len(s5.vertices()._selection) == 2
    assert s5.vertices(">Y")._selection[0].Center().y == approx(0)

    s6 = Sketch().arc((0, 0), 1, 90, 360)

    assert len(s6.vertices()._selection) == 1


def test_bezier():
    s1 = (
        Sketch()
        .segment((0, 0), (0, 0.5))
        .bezier(((0, 0.5), (-1, 2), (1, 0.5), (5, 0)))
        .bezier(((5, 0), (1, -0.5), (-1, -2), (0, -0.5)))
        .close()
        .assemble()
    )
    assert s1._faces.Area() == approx(5.35)
    # What other kind of tests can we do?


def test_assemble():

    s1 = Sketch()
    s1.segment((0.0, 0), (0.0, 2.0))
    s1.segment(Vector(4.0, -1)).close().arc((0.7, 0.6), 0.4, 0.0, 360.0).assemble()

    s2 = Sketch()
    s2.segment((0, 0), (1, 0))
    s2.segment((2, 0), (3, 0))
    with raises(ValueError):
        s2.assemble()


def test_finalize():

    parent = object()
    s = Sketch(parent).rect(2, 2).circle(0.5, mode="s")

    assert s.finalize() is parent


def test_misc():

    with raises(ValueError):
        Sketch()._startPoint()

    with raises(ValueError):
        Sketch()._endPoint()


def test_located():

    s1 = Sketch().segment((0, 0), (1, 0)).segment((1, 1)).close().assemble()

    assert len(s1._edges) == 3
    assert len(s1._faces.Faces()) == 1

    s2 = s1.located(loc=Location())

    assert len(s2._edges) == 0
    assert len(s2._faces.Faces()) == 1


def test_constraint_validation():

    with raises(ValueError):
        Sketch().segment(1.0, 1.0, "s").constrain("s", "Dummy", None)

    with raises(ValueError):
        Sketch().segment(1.0, 1.0, "s").constrain("s", "s", "Fixed", None)

    with raises(ValueError):
        Sketch().spline([(1.0, 1.0), (2.0, 1.0), (0.0, 0.0)], "s").constrain(
            "s", "Fixed", None
        )

    with raises(ValueError):
        Sketch().segment(1.0, 1.0, "s").constrain("s", "Fixed", 1)


def test_constraint_solver():

    s1 = (
        Sketch()
        .segment((0.0, 0), (0.0, 2.0), "segment1")
        .segment((0.5, 2.5), (1.0, 1), "segment2")
        .close("segment3")
    )
    s1.constrain("segment1", "Fixed", None)
    s1.constrain("segment1", "segment2", "Coincident", None)
    s1.constrain("segment2", "segment3", "Coincident", None)
    s1.constrain("segment3", "segment1", "Coincident", None)
    s1.constrain("segment3", "segment1", "Angle", 90)
    s1.constrain("segment2", "segment3", "Angle", 180 - 45)

    s1.solve()

    assert s1._solve_status["status"] == 4

    s1.assemble()

    assert s1._faces.isValid()

    s2 = (
        Sketch()
        .arc((0.0, 0.0), (-0.5, 0.5), (0.0, 1.0), "arc1")
        .arc((0.0, 1.0), (0.5, 1.5), (1.0, 1.0), "arc2")
        .segment((1.0, 0.0), "segment1")
        .close("segment2")
    )

    s2.constrain("segment2", "Fixed", None)
    s2.constrain("segment1", "segment2", "Coincident", None)
    s2.constrain("arc2", "segment1", "Coincident", None)
    s2.constrain("segment2", "arc1", "Coincident", None)
    s2.constrain("arc1", "arc2", "Coincident", None)
    s2.constrain("segment1", "segment2", "Angle", 90)
    s2.constrain("segment2", "arc1", "Angle", 90)
    s2.constrain("arc1", "arc2", "Angle", -90)
    s2.constrain("arc2", "segment1", "Angle", 90)
    s2.constrain("segment1", "Length", 0.5)
    s2.constrain("arc1", "Length", 1.0)

    s2.solve()

    assert s2._solve_status["status"] == 4

    s2.assemble()

    assert s2._faces.isValid()

    assert s2._tags["segment1"][0].Length() == approx(0.5)
    assert s2._tags["arc1"][0].Length() == approx(1.0)

    s3 = (
        Sketch()
        .arc((0.0, 0.0), (-0.5, 0.5), (0.0, 1.0), "arc1")
        .segment((1.0, 0.0), "segment1")
        .close("segment2")
    )

    s3.constrain("segment2", "Fixed", None)
    s3.constrain("arc1", "ArcAngle", 60)
    s3.constrain("arc1", "Radius", 1.0)
    s3.constrain("segment2", "arc1", "Coincident", None)
    s3.constrain("arc1", "segment1", "Coincident", None)
    s3.constrain("segment1", "segment2", "Coincident", None)

    s3.solve()

    assert s3._solve_status["status"] == 4

    s3.assemble()

    assert s3._faces.isValid()

    assert s3._tags["arc1"][0].radius() == approx(1)
    assert s3._tags["arc1"][0].Length() == approx(pi / 3)

    s4 = (
        Sketch()
        .arc((0.0, 0.0), (-0.5, 0.5), (0.0, 1.0), "arc1")
        .segment((1.0, 0.0), "segment1")
        .close("segment2")
    )

    s4.constrain("segment2", "Fixed", None)
    s4.constrain("segment1", "Orientation", (-1.0, -1))
    s4.constrain("segment1", "segment2", "Distance", (0.0, 0.5, 2.0))
    s4.constrain("segment2", "arc1", "Coincident", None)
    s4.constrain("arc1", "segment1", "Coincident", None)
    s4.constrain("segment1", "segment2", "Coincident", None)

    s4.solve()

    assert s4._solve_status["status"] == 4

    s4.assemble()

    assert s4._faces.isValid()

    seg1 = s4._tags["segment1"][0]
    seg2 = s4._tags["segment2"][0]

    assert (seg1.endPoint() - seg1.startPoint()).getAngle(Vector(-1, -1)) == approx(
        0, abs=1e-9
    )

    midpoint = (seg2.startPoint() + seg2.endPoint()) / 2

    assert (midpoint - seg1.startPoint()).Length == approx(2)

    s5 = (
        Sketch()
        .segment((0, 0), (0, 3.0), "segment1")
        .arc((0.0, 0), (1.5, 1.5), (0.0, 3), "arc1")
        .arc((0.0, 0), (-1.0, 1.5), (0.0, 3), "arc2")
    )

    s5.constrain("segment1", "Fixed", None)
    s5.constrain("segment1", "arc1", "Distance", (0.5, 0.5, 3))
    s5.constrain("segment1", "arc1", "Distance", (0.0, 1.0, 0.0))
    s5.constrain("arc1", "segment1", "Distance", (0.0, 1.0, 0.0))
    s5.constrain("segment1", "arc2", "Coincident", None)
    s5.constrain("arc2", "segment1", "Coincident", None)
    s5.constrain("arc1", "arc2", "Distance", (0.5, 0.5, 10.5))

    s5.solve()

    assert s5._solve_status["status"] == 4

    mid0 = s5._edges[0].positionAt(0.5)
    mid1 = s5._edges[1].positionAt(0.5)
    mid2 = s5._edges[2].positionAt(0.5)

    assert (mid1 - mid0).Length == approx(3)
    assert (mid1 - mid2).Length == approx(10.5)

    s6 = (
        Sketch()
        .segment((0, 0), (0, 3.0), "segment1")
        .arc((0.0, 0), (5.5, 5.5), (0.0, 3), "arc1")
    )

    s6.constrain("segment1", "Fixed", None)
    s6.constrain("segment1", "arc1", "Coincident", None)
    s6.constrain("arc1", "segment1", "Coincident", None)
    s6.constrain("arc1", "segment1", "Distance", (None, 0.5, 0))

    s6.solve()

    assert s6._solve_status["status"] == 4

    mid0 = s6._edges[0].positionAt(0.5)
    mid1 = s6._edges[1].positionAt(0.5)

    assert (mid1 - mid0).Length == approx(1.5)

    s7 = (
        Sketch()
        .segment((0, 0), (0, 3.0), "segment1")
        .arc((0.0, 0), (5.5, 5.5), (0.0, 4), "arc1")
    )

    s7.constrain("segment1", "FixedPoint", 0)
    s7.constrain("arc1", "FixedPoint", None)
    s7.constrain("arc1", "FixedPoint", 1)
    s7.constrain("arc1", "segment1", "Distance", (0, 0, 0))
    s7.constrain("arc1", "segment1", "Distance", (1, 1, 0))

    s7.solve()

    assert s7._solve_status["status"] == 4

    s7.assemble()

    assert s7._faces.isValid()


def test_dxf_import():

    filename = os.path.join(testdataDir, "gear.dxf")

    s1 = Sketch().importDXF(filename, tol=1e-3)

    assert s1._faces.isValid()

    s2 = Sketch().importDXF(filename, tol=1e-3).circle(5, mode="s")

    assert s2._faces.isValid()

    s3 = Sketch().circle(20).importDXF(filename, tol=1e-3, mode="s")

    assert s3._faces.isValid()

    s4 = Sketch().importDXF(filename, tol=1e-3, include=["0"])

    assert s4._faces.isValid()

    s5 = Sketch().importDXF(filename, tol=1e-3, exclude=["1"])

    assert s5._faces.isValid()


def test_val():

    s1 = Sketch().segment((0, 0), (0, 1))

    assert isinstance(s1.val(), Edge)

    s1.vertices()

    assert isinstance(s1.val(), Vertex)

    s2 = Sketch().circle(1)

    assert len(s2.val().Faces()) == 1


def test_vals():

    s1 = Sketch().segment((0, 0), (0, 1))

    assert len(s1.vals()) == 1

    s1.vertices()

    assert len(s1.vals()) == 2

    s2 = Sketch().circle(1)

    assert len(s2.vals()) == 1


def test_bool_ops():

    s1 = Sketch().rect(0.5, 2)
    s2 = Sketch().rect(1, 1)
    s3 = Sketch().segment((-1, 0), (1, 0))

    # sum
    assert (s1 + s2).val().Area() == approx(1.5)
    # diff
    assert (s1 - s2).val().Area() == approx(0.5)
    assert len((s1 - s2).val().Wires()) == 2
    # common
    assert (s1 * s2).val().Area() == approx(0.5)
    assert len((s1 * s2).val().Wires()) == 1
    # split
    assert len((s1 / s3).val().Faces()) == 2
    assert len((s1 / s2).val().Faces()) == 3
    assert (s1 / s2).val().Area() == approx(1)


def test_export():

    s1 = Sketch().rect(1, 1).export("sketch.dxf")
    s2 = Sketch().importDXF("sketch.dxf")

    assert (s1 - s2).val().Area() == approx(0)


@fixture
def s1():
    """
    Sketch with 2 faces

    """

    return Sketch().rect(1, 1).push([(5, 0)]).rect(1, 4).reset()


@fixture
def s2():
    """
    Sketch with 5 faces

    """

    return Sketch().rarray(2, 1, 5, 1).rect(1, 1).reset()


@fixture
def f():
    """
    Single face
    """

    return Sketch().rect(0.1, 1).val()


def test_filter(s1, f):

    assert len(s1.filter(lambda x: x.Area() > 2).vals()) == 1
    assert len(s1.reset().filter(lambda x: x.Area() >= 1).vals()) == 2
    assert len(s1.filter(lambda x: x.Area() < 1).vals()) == 0


def test_sort(s1):

    assert s1.sort(lambda x: -x.Area())[-1].val().Area() == approx(1)


def test_apply(s1, f):

    assert s1.apply(lambda x: [f]).val().Area() == approx(0.1)


def test_map(s1, f):
    assert s1.map(
        lambda x: f.moved(x.Center())
    ).replace().reset().val().Area() == approx(0.2)


def test_getitem(s2):

    assert len(s2[0].vals()) == 1
    assert len(s2.reset()[-2:].vals()) == 2
    assert len(s2.reset()[[0, 1]].vals()) == 2


def test_invoke(s1, s2):

    # builtin
    assert len(s2.invoke(print).vals()) == 5
    # arity 0
    assert len(s2.invoke(lambda: 1).vals()) == 5
    # arity 1 and no return
    assert len(s2.invoke(lambda x: None).vals()) == 5
    # arity 1
    assert len(s2.invoke(lambda x: s1).vals()) == 2
    # test exception with wrong arity
    with raises(ValueError):
        s2.invoke(lambda x, y: 1)


@fixture
def s3():
    """
    Simple sketch with one face.
    """

    return Sketch().rect(1, 1)


def test_replace(s3, f):

    assert s3.map(lambda x: f).replace().reset().val().Area() == approx(0.1)


def test_add(s3, f):

    # we do not clean, so adding will increase the number of edges
    assert len(s3.map(lambda x: f).add().reset().val().Edges()) == 10


def test_subtract(s3, f):

    assert s3.map(lambda x: f).subtract().reset().val().Area() == approx(0.9)


def test_iter(s1):

    # __iter__ ofer face
    assert len(list(s1)) == 2

    # __iter__ over edges
    assert len(list(Sketch().segment((0, 0), (1, 0)))) == 1


def test_sanitize():

    # it does not make sense to fuse faces and Locations
    with raises(ValueError):
        Sketch().rect(1, 1) + Sketch().rarray(1, 1, 1, 1)


def test_missing_selection(s1):

    # offset requires selected wires
    with raises(ValueError):
        s1.offset(0.1)

    # fillet requires selected vertices
    with raises(ValueError):
        s1.fillet(0.1)

    # cannot delete without selection
    with raises(ValueError):
        s1.delete()

    # cannot tag without selection
    with raises(ValueError):
        s1.tag("name")
