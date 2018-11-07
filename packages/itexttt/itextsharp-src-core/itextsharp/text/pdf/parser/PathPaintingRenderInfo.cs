﻿using System;
using System.Collections.Generic;
using System.Text;

namespace iTextSharp.text.pdf.parser {

    /**
     * Contains information relating to painting current path.
     *
     * @since 5.5.6
     */
    public class PathPaintingRenderInfo {

        /**
         * The nonzero winding number rule determines whether a given point is inside a path by
         * conceptually drawing a ray from that point to infinity in any direction and then examining
         * the places where a segment of the path crosses the ray. Starting with a count of 0, the rule
         * adds 1 each time a path segment crosses the ray from left to right and subtracts 1 each time a
         * segment crosses from right to left. After counting all the crossings, if the result is 0, the
         * point is outside the path; otherwise, it is inside.
         *
         * For more details see PDF spec.
         */
        public const int NONZERO_WINDING_RULE = 1;

        /**
         * The even-odd rule determines whether a point is inside a path by drawing a ray from that point in
         * any direction and simply counting the number of path segments that cross the ray, regardless of
         * direction. If this number is odd, the point is inside; if even, the point is outside.
         *
         * For more details see PDF spec.
         */
        public const int EVEN_ODD_RULE = 2;

        /**
         * End the path object without filling or stroking it. This operator shall be a path-painting no-op,
         * used primarily for the side effect of changing the current clipping path
         */
        public const int NO_OP = 0;

        /**
         * Value specifying stroke operation to perform on the current path.
         */
        public const int STROKE = 1;

        /**
         * Value specifying fill operation to perform on the current path. When the fill operation
         * is performed it should use either nonzero winding or even-odd rule.
         */
        public const int FILL = 2;

        private int operation;
        private int rule;
        private GraphicsState gs;

        /**
         * @param operation One of the possible combinations of {@link #STROKE} and {@link #FILL} values or {@link #NO_OP}
         * @param rule      Either {@link #NONZERO_WINDING_RULE} or {@link #EVEN_ODD_RULE}.
         * @param gs        The graphics state.
         */
        public PathPaintingRenderInfo(int operation, int rule, GraphicsState gs) {
            this.operation = operation;
            this.rule = rule;
            this.gs = gs;
        }

        /**
         * If the operation is {@link #NO_OP} then the rule is ignored,
         * otherwise {@link #NONZERO_WINDING_RULE} is used by default.
         *
         * See {@link #PathPaintingRenderInfo(int, int, GraphicsState)}
         */
        public PathPaintingRenderInfo(int operation, GraphicsState gs) : this(operation, NONZERO_WINDING_RULE, gs) {
        }

        /**
         * @return <CODE>int</CODE> value which is either {@link #NO_OP} or one of possible
         * combinations of {@link #STROKE} and {@link #FILL}
         */
        public virtual int Operation {
            get { return operation; }
        }

        /**
         * @return Either {@link #NONZERO_WINDING_RULE} or {@link #EVEN_ODD_RULE}.
         */
        public virtual int Rule {
            get { return rule; }
        }

        /**
         * @return Current transformation matrix.
         */
        public virtual Matrix Ctm {
            get { return gs.ctm; }
        }

        public virtual float LineWidth {
            get { return gs.LineWidth; }
        }

        public virtual int LineCapStyle {
            get { return gs.LineCapStyle; }
        }

        public virtual int LineJoinStyle {
            get { return gs.LineJoinStyle; }
        }

        public virtual float MiterLimit {
            get { return gs.MiterLimit; }
        }

        public virtual LineDashPattern LineDashPattern {
            get { return gs.GetLineDashPattern(); }
        }
    }
}
